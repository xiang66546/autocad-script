# -*- coding: utf-8 -*-
"""
Created on Sun Apr 13 11:12:14 2025

@author: Administrator
"""

import math
import win32com.client
from win32com.client import VARIANT
import pandas as pd
import time

from collections import defaultdict
from datetime import datetime
from scipy.spatial import KDTree
import numpy as np
import networkx as nx



#%% class 讀取目標圖層資料 -> geometry_dict
class CADGeometryExtractor:
    def __init__(self, doc, target_layer):
        self.doc = doc
        self.model_space = self.doc.ModelSpace
        self.target_layer = target_layer
        self.geometry_dict = {"LINE": {}, "ARC": {}} 
        self.handle_to_object = {}
        self.tolerance = 1e-6  # 容差設定，保護浮點數運算精度

    def points_close(self, p1, p2):
        return math.dist(p1, p2) < self.tolerance

    def normalize_segment(self, start, end):
        return tuple(sorted([start, end]))

    def _check_endpoints_exist(self, start, end):
        """
        檢查給定的 StartPoint 和 EndPoint 是否已存在於 geometry_dict 中。
        start, end: 2D 座標 tuple (x, y)
        返回 True 如果端點已存在，否則返回 False。
        """
        for line_data in self.geometry_dict["LINE"].values():
            existing_start = line_data["StartPoint"]
            existing_end = line_data["EndPoint"]
            if ((self.points_close(start, existing_start) and self.points_close(end, existing_end)) or
                (self.points_close(start, existing_end) and self.points_close(end, existing_start))):
                return True

        for arc_data in self.geometry_dict["ARC"].values():
            existing_start = arc_data["StartPoint"]
            existing_end = arc_data["EndPoint"]
            if ((self.points_close(start, existing_start) and self.points_close(end, existing_end)) or
                (self.points_close(start, existing_end) and self.points_close(end, existing_start))):
                return True

        return False

    def extract(self):
        for entity in self.model_space:
            if entity.Layer != self.target_layer:
                continue

            obj_name = entity.ObjectName

            if obj_name == 'AcDbLine':
                self._extract_line(entity)

            elif obj_name == 'AcDbArc':
                self._extract_arc(entity)

            elif obj_name in ('AcDbPolyline', 'AcDb2dPolyline'):
                self._explode_polyline(entity)  # 無論是否閉合，都進行分解

    def _extract_line(self, line):
        handle = line.Handle
        start = tuple(line.StartPoint[:2])
        end = tuple(line.EndPoint[:2])

        # 檢查端點是否已存在
        if self._check_endpoints_exist(start, end):
            return

        delta = (end[0] - start[0], end[1] - start[1])
        angle = math.degrees(math.atan2(delta[1], delta[0])) % 360

        self.geometry_dict["LINE"][handle] = {
            "StartPoint": start,
            "EndPoint": end,
            "Length": math.dist(start, end),
            "Angle": angle,
            "Delta": delta
        }
        self.handle_to_object[handle] = line

    def _extract_arc(self, arc):
        handle = arc.Handle
        center = tuple(arc.Center[:2])
        start = tuple(arc.StartPoint[:2])
        end = tuple(arc.EndPoint[:2])

        # 檢查端點是否已存在
        if self._check_endpoints_exist(start, end):
            return

        # 計算凸度 (Bulge)
        start_angle = arc.StartAngle
        end_angle = arc.EndAngle
        included_angle = end_angle - start_angle
        # 正規化包含角，確保為正值
        if included_angle < 0:
            included_angle += 2 * math.pi
        # 計算凸度：tan(包含角/4)
        bulge = math.tan(included_angle / 4)

        self.geometry_dict["ARC"][handle] = {
            "Center": center,
            "Radius": arc.Radius,
            "StartAngle": arc.StartAngle,
            "EndAngle": arc.EndAngle,
            "StartPoint": start,
            "EndPoint": end,
            "Bulge": bulge
        }
        self.handle_to_object[handle] = arc

    def _explode_polyline(self, pline):
        """分解 Polyline（無論是否閉合）"""
        try:
            exploded_items = pline.Explode()  # 爆炸成多個元素

            for item in exploded_items:
                if item.ObjectName == 'AcDbLine':
                    self._extract_line(item)
                elif item.ObjectName == 'AcDbArc':
                    self._extract_arc(item)
                else:
                    print(f"⚠️ 爆炸後發現奇怪物件: {item.ObjectName}，略過。")
        except Exception as e:
            print(f"⚠️ 爆炸Polyline失敗 (Handle={pline.Handle})，錯誤: {e}")

    def get_geometry(self):
        return self.geometry_dict

    def get_object_mapping(self):
        return self.handle_to_object

    def get_object_by_handle(self, handle):
        return self.handle_to_object.get(handle, None)


#%% main

start_time = time.time()  # ⏱️ 開始計時

try:
    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    model_space = doc.ModelSpace
    print(f"✅ 已連接到文件: {doc.Name}")
except Exception as e:
    print(f"Error: Unable to connect to AutoCAD: {e}")
    exit()

# target_layers = "674_分區用地界線"   
target_layers = "001-街廓"  
# target_layers = "菓林人行道$0$01-街廓線"  
# target_layers = "0-細部計畫線"  
#20160808 細部計畫$0$0-分區街廓(都發局0707)


# 🔍 檢查圖層是否存在
layer_names = [layer.Name for layer in doc.Layers]
if target_layers not in layer_names:
    print(f"❌ 圖層 '{target_layers}' 不存在於此圖檔中！")
    print(f"📋 可用圖層清單：{layer_names}")
    exit()
    

print("🔎 提取圖層資料中...")
extractor = CADGeometryExtractor(doc, target_layers)
extractor.extract()
geometry_dict = extractor.get_geometry()
object_mapping = extractor.get_object_mapping()

end_time = time.time()  # ⏱️ 結束計時
print(f"幾何提取完成，用時：{end_time - start_time:.2f} 秒")


    
#%% 將polyline繪製到cad
# 計算距離
def calculate_distance(point_1, point_2):
    x1, y1 = point_1
    x2, y2 = point_2
    return ((x1 - x2) ** 2 + (y1 - y2) ** 2) ** 0.5


# 構建圖（獲取每一個handle可以連接的其他handle）
def build_graph_kdtree_numpy(coor_df, threshold):
    graph = defaultdict(list)
    handles = np.array(coor_df.index)

    # 將所有起點與終點轉為 numpy array
    coordinates = np.array([coor_df.loc[h, 'Coordinate'] for h in handles], dtype=object)
    start_points = np.array([coord[0] for coord in coordinates])
    end_points = np.array([coord[1] for coord in coordinates])

    # 建立 handle -> 去掉 _r 對應表，加速比較
    original_handles = np.array([h.replace('_r', '') for h in handles])

    # 用起點建 KDTree，讓終點去查
    tree = KDTree(start_points)

    for i, handle1 in enumerate(handles):
        end_point = end_points[i]
        idx_list = tree.query_ball_point(end_point, threshold)

        for j in idx_list:
            handle2 = handles[j]

            if handle1 == handle2:
                continue

            if original_handles[i] == original_handles[j]:  # 是反向版本就略過
                continue

            graph[handle1].append(handle2)

    return graph


# 獲取路徑的頂點座標
def get_path_points(path, coor_df):
    points = []
    for i, handle in enumerate(path):
        coords = coor_df.loc[handle, 'Coordinate']
        start, end = coords

        if i == 0:
            points.append(start)
        points.append(end)

    if calculate_distance(points[0], points[-1]) < threshold:
        return points
    return points + [points[0]]

def deduplicate_cycles(cycles):
    seen = set()
    deduped = []

    for cycle in cycles:
        normalized = frozenset(h.replace('_r', '') for h in cycle)

        if normalized not in seen:
            seen.add(normalized)
            deduped.append(cycle)

    return deduped




def get_coord_df(geometry_dict):
    coor_df = pd.DataFrame(columns=['Coordinate', 'Bulge'])
    seen_handles = set()
    for geom_type in ['LINE', 'ARC']:
        for handle, obj_data in geometry_dict[geom_type].items():
            if handle in seen_handles:
                print(f"⚠️ handle {handle} 已存在，跳過")
                continue
            seen_handles.add(handle)

            start_point = obj_data['StartPoint']
            end_point = obj_data['EndPoint']
            bulge = 0 if geom_type == 'LINE' else obj_data['Bulge']

            # 正向記錄
            row = pd.DataFrame({
                'Coordinate': [(start_point, end_point)],
                'Bulge': [bulge]
            }, index=[handle])
            coor_df = pd.concat([coor_df, row])

            # ✅ 只有 problem_handles 才建立反向版本
            if handle in problem_handles:
                reversed_handle = f"{handle}_r"
                reversed_bulge = -bulge if geom_type == 'ARC' else bulge
                row_reversed = pd.DataFrame({
                    'Coordinate': [(end_point, start_point)],
                    'Bulge': [reversed_bulge]
                }, index=[reversed_handle])
                coor_df = pd.concat([coor_df, row_reversed])
                
    return coor_df


def classify_connection_types_from_geometry_dict(geometry_dict, threshold):
    # 合併 LINE 和 ARC handle 的起點與終點
    all_segments = {}
    for geom_type in ['LINE', 'ARC']:
        for handle, data in geometry_dict[geom_type].items():
            start = data['StartPoint']
            end = data['EndPoint']
            all_segments[handle] = (start, end)

    handles = list(all_segments.keys())
    results = {
        'end_to_start': [],
        'end_to_end': [],
        'start_to_start': [],
        'start_to_end': []
    }

    for h1 in handles:
        s1, e1 = all_segments[h1]

        for h2 in handles:
            if h1 == h2:
                continue

            s2, e2 = all_segments[h2]

            if calculate_distance(e1, s2) < threshold:
                results['end_to_start'].append((h1, h2))
            elif calculate_distance(e1, e2) < threshold:
                results['end_to_end'].append((h1, h2))
            elif calculate_distance(s1, s2) < threshold:
                results['start_to_start'].append((h1, h2))
            elif calculate_distance(s1, e2) < threshold:
                results['start_to_end'].append((h1, h2))
                
                
    problem_handles = list(set(h for h, _ in results['end_to_end'] + results['start_to_start']))
    return problem_handles



def filter_outer_cycles_by_geometry(cycles, handle_coords, buffer_tolerance=0.01):
    from shapely.geometry import Polygon

    poly_with_area = []
    for cycle in cycles:
        points = []
        for i, h in enumerate(cycle):
            start, end = handle_coords[h]
            if i == 0:
                points.append(start)
            points.append(end)
        polygon = Polygon(points)
        if polygon.is_valid and polygon.area > 0:
            poly_with_area.append((cycle, polygon))

    keep = []
    for i, (cycle_i, poly_i) in enumerate(poly_with_area):
        is_inner = False
        for j, (cycle_j, poly_j) in enumerate(poly_with_area):
            if i == j:
                continue
            # 使用 buffer 容差判定是否被包在別人裡面
            if poly_i.within(poly_j.buffer(buffer_tolerance)):
                is_inner = True
                break
        if not is_inner:
            keep.append(cycle_i)

    return keep



def get_polyline_path_list(coor_df, threshold):
    s = time.time()
    
    graph = build_graph_kdtree_numpy(coor_df, threshold)
    
    G = nx.DiGraph(graph)
    cycles = list(nx.simple_cycles(G))
    # cycles = deduplicate_cycles(cycles)
    
    handle_coords = {h: coor_df.at[h, 'Coordinate'] for h in coor_df.index}
    
    # 🎯 篩選掉內部的封閉區域
    outer_cycles = filter_outer_cycles_by_geometry(cycles, handle_coords, buffer_tolerance=0.01)    

    e = time.time()
    print(f"完成，用時：{e - s:.2f} 秒")
    return outer_cycles


def draw_polyline(layer_name,polyline_path_list,  coor_df):
    try:
        layer_obj = doc.Layers.Item(layer_name)
    except:
        layer_obj = doc.Layers.Add(layer_name)
    layer_obj.color = 1
    
    for polyline_handles in polyline_path_list:
        vertices = []
        bulges = []
        
        points = get_path_points(polyline_handles, coor_df)
    
        for i, handle in enumerate(polyline_handles):
            # 移除 _r 後綴以映射回原始 handle
            original_handle = handle.replace('_r', '')
            geom_type = 'LINE' if original_handle in geometry_dict['LINE'] else 'ARC'
            obj_data = geometry_dict[geom_type][original_handle]
            # 使用 coor_df 中的座標，而不是 geometry_dict，因為方向可能已反轉
            start_point, end_point = coor_df.loc[handle, 'Coordinate']
            bulge = coor_df.loc[handle, 'Bulge']
    
            if i == 0:
                vertices.extend([start_point[0], start_point[1]])
                bulges.append(bulge)
    
            vertices.extend([end_point[0], end_point[1]])
            if i < len(polyline_handles) - 1:
                next_handle = polyline_handles[i + 1]
                next_bulge = coor_df.loc[next_handle, 'Bulge']
                bulges.append(next_bulge)
            else:
                bulges.append(0)
    
        vertices_array = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, vertices)
        polyline = model_space.AddLightWeightPolyline(vertices_array)
        polyline.Layer = layer_name
        
        for i in range(len(bulges)):
            polyline.SetBulge(i, bulges[i])
    
        polyline.Update()
        

print("🔎 建立polyline中...")
threshold = 0.1

#座標會反向連接，因此將end_to_end跟start_to_start的座標在coord_df中新增一個反向
problem_handles = classify_connection_types_from_geometry_dict(geometry_dict, threshold)

#取得handle座標df
coor_df = get_coord_df(geometry_dict)

#取得最終polyline的list
polyline_path_list = get_polyline_path_list(coor_df, threshold)

# 繪製到 AutoCAD
now_str = datetime.now().strftime('%Y%m%d%H%M%S')
layer_name = f"test{now_str}"
draw_polyline(layer_name,polyline_path_list,  coor_df)

    
#%% get polyline


'''
因為前面有可能沒有全部街廓都抓到
因此需要手動修正後再抓一樣的圖層
'''


def extract_polylines_from_layer(doc, layer_name):
    """
    從指定圖層讀取所有 Polyline（AcDbPolyline / AcDb2dPolyline）資料。
    回傳每個 polyline 的：
        - handle
        - points: List of (x, y)
        - bulges: List of bulge
        - closed: 是否閉合
    """
    polylines = []
    model_space = doc.ModelSpace

    for entity in model_space:
        if entity.Layer != layer_name:
            continue
        if entity.ObjectName not in ['AcDbPolyline', 'AcDb2dPolyline']:
            continue

        try:
            coords = list(entity.Coordinates)  # 平面點座標：x1,y1,x2,y2,...
            points = [(coords[i], coords[i+1]) for i in range(0, len(coords), 2)]

            bulges = []
            for i in range(len(points) - 1):  # 最後一段的 bulge 通常預設 0
                try:
                    bulges.append(entity.GetBulge(i))
                except:
                    bulges.append(0)
            bulges.append(0)  # 最後一段沒有下一段可接，bulge 補 0

            polylines.append({
                'handle': entity.Handle,
                'points': points,
                'bulges': bulges,
                'closed': entity.Closed
            })

        except Exception as e:
            print(f"⚠️ 無法處理 handle {entity.Handle}，原因：{e}")
            continue

    return polylines



# layer_name = 'test20250429215447'
layer_name = 'test20250502155511'
print("🔎 讀取polyline中...")
polylines = extract_polylines_from_layer(doc, layer_name)



#%% 畫角平分線

import math
import pythoncom
from win32com.client import VARIANT
from math import hypot
import heapq

#### ── 向量與角度工具 ────────────────────────────────────────────────

def angle_between(v1, v2):
    """
    計算兩個向量 v1, v2 之間的夾角（0°–180°）。
    v1, v2 都是 (x, y) tuple 或 list。
    """
    dot   = v1[0]*v2[0] + v1[1]*v2[1]
    cross = v1[0]*v2[1] - v1[1]*v2[0]
    theta = math.atan2(abs(cross), dot)
    return math.degrees(theta)

def is_clockwise(points):
    # 判斷多邊形頂點是否按順時針排列
    total = 0
    for i in range(len(points)-1):
        x1,y1 = points[i]
        x2,y2 = points[i+1]
        total += (x2-x1)*(y2+y1)
    return total > 0

def get_tangent_at_point(start_point, end_point, bulge, at_start=True):
    """
    計算圓弧在端點的單位切線向量。
    """
    chord = math.hypot(end_point[0]-start_point[0], end_point[1]-start_point[1])
    if chord < 1e-12 or abs(bulge) < 1e-6:
        # 視為直線：起點切線方向=弦向量，終點切線方向=弦向量相反
        vx,vy = end_point[0]-start_point[0], end_point[1]-start_point[1]
        mag = math.hypot(vx, vy)
        if mag<1e-12: return (0,0)
        if at_start:
            return (vx/mag, vy/mag)
        else:
            return (-vx/mag, -vy/mag)

    # Bulge = tan(theta/4)
    theta = 4 * math.atan(abs(bulge))
    radius = chord / (2*math.sin(theta/2))
    mid_x = (start_point[0]+end_point[0])/2
    mid_y = (start_point[1]+end_point[1])/2
    dx,dy  = end_point[0]-start_point[0], end_point[1]-start_point[1]
    # 法向量方向取決於 bulge 正負
    nx_,ny_ = (-dy, dx) if bulge>0 else (dy, -dx)
    nmag = math.hypot(nx_, ny_)
    nx_,ny_ = nx_/nmag, ny_/nmag
    h = radius * math.cos(theta/2)
    cx,cy = mid_x + nx_*h, mid_y + ny_*h  # 圓心
    point = start_point if at_start else end_point
    rx,ry = point[0]-cx, point[1]-cy
    # 切線向量
    tx,ty = (-ry, rx) if bulge>0 else (ry, -rx)
    tmag = math.hypot(tx,ty)
    return (tx/tmag, ty/tmag) if tmag>1e-12 else (0,0)

def get_dynamic_bisector_length(points):
    # 多邊形邊界框對角線 5%
    xs = [p[0] for p in points]; ys = [p[1] for p in points]
    diag = math.hypot(max(xs)-min(xs), max(ys)-min(ys))
    return max(diag*0.05, 3)

def calculate_angle_bisector(p1, p2, p3, bulge1, bulge2, points):
    """
    計算頂點 p2 處，p1-p2-p3 三點形成之角的角平分線，回傳 [起點, 終點]。
    """
    L = get_dynamic_bisector_length(points)
    # 向量 v1 從 p2 指向 p1
    if abs(bulge1)<1e-6:
        v1 = (p1[0]-p2[0], p1[1]-p2[1])
    else:
        t1 = get_tangent_at_point(p1, p2, bulge1, at_start=False)
        v1 = (-t1[0], -t1[1])
    # 向量 v2 從 p2 指向 p3
    if abs(bulge2)<1e-6:
        v2 = (p3[0]-p2[0], p3[1]-p2[1])
    else:
        v2 = get_tangent_at_point(p2, p3, bulge2, at_start=True)

    m1 = math.hypot(*v1); m2 = math.hypot(*v2)
    if m1<1e-12 or m2<1e-12:
        return [p2, p2]
    u1, u2 = (v1[0]/m1, v1[1]/m1), (v2[0]/m2, v2[1]/m2)
    bis = (u1[0]+u2[0], u1[1]+u2[1])
    bmag = math.hypot(*bis)
    if bmag<1e-12:
        return [p2, p2]
    bis = (bis[0]/bmag, bis[1]/bmag)

    # 確保指向內部
    test = (p2[0]+bis[0]*L*0.5, p2[1]+bis[1]*L*0.5)
    try:
        from shapely.geometry import Polygon, Point
        poly = Polygon(points)
        if poly.is_valid and not poly.contains(Point(test)):
            bis = (-bis[0], -bis[1])
    except:
        # fallback：使用繞行方向
        cw = is_clockwise(points)
        cross = u1[0]*u2[1] - u1[1]*u2[0]
        if (cross<0 and not cw) or (cross>0 and cw):
            bis = (-bis[0], -bis[1])

    endpt = (p2[0]+bis[0]*L, p2[1]+bis[1]*L)
    return [p2, endpt]

#### ── 角度序列、分組函式 ─────────────────────────────────────────
def dist(p1, p2):
    """計算兩點之間的距離"""
    return math.hypot(p2[0] - p1[0], p2[1] - p1[1])


def find_consecutive_angle_runs(entries):
    angles = [entry[-1] for entry in entries]
    all_same_angle = all(angle == angles[0] for angle in angles)
    
    if all_same_angle:
        n = len(entries)
        runs = []
        
        # 計算 entries[0] 與 entries[-1] 以及 entries[1] 的距離
        curr_run = entries[0]
        prev_run = entries[-1]
        next_run = entries[1]
        
        prev_distance = dist(curr_run[1], prev_run[1])
        next_distance = dist(curr_run[1], next_run[1])
        
        if prev_distance < next_distance:
            # entries[0] 與 entries[-1] 配對
            runs.append([entries[-1], entries[0]])
            start, end = 1, n-1
        else:
            start, end = 0, n
            
        for i in range(start, end, 2):
            if i + 1 < end:
                runs.append([entries[i], entries[i+1]])
            else:
                runs.append([entries[i]])  # 單獨成段
   
        return runs 
    
    else:
        runs = []
        cur = [entries[0]]
        for prev, curr in zip(entries, entries[1:]):
            spatial_dist = dist(curr[1], prev[1])
            angle_diff = abs(curr[-1] - prev[-1])
            if spatial_dist < 20 and angle_diff < 0.05:
                cur.append(curr)
            else:
                runs.append(cur); cur=[curr]
        runs.append(cur)
        # 合併首尾同角度
        if len(runs)>1 and runs[0][0][-1]==runs[-1][0][-1] or abs(runs[0][0][-1] - runs[-1][0][-1]) < 0.05:
            runs[0] = runs[-1]+runs[0]
            runs.pop()
            
        i = 0
        while i < len(runs):
            run = runs[i]
            # 只處理「只有一個 entry 的那種 run」
            if len(run) == 1 and len(runs) > 1:
                angle = run[0][-1]
                prev_i = (i - 1) % len(runs)
                next_i = (i + 1) % len(runs)
                prev_angle = runs[prev_i][0][-1]
                next_angle = runs[next_i][0][-1]
        
                # 如果跟前一組角度差小於 12，就合併到前一組
                if abs(angle - prev_angle) < 12:
                    runs[prev_i].extend(run)
                    runs.pop(i)
                    # 如果我們合併的是第一組，i 不動，否則回退一格以檢查新合併後的前一組
                    if i > prev_i:
                        i -= 1
                    continue
        
                # 否則如果跟下一組角度差小於 12，就合併到下一組
                if abs(angle - next_angle) < 12:
                    runs[next_i] = run + runs[next_i]
                    runs.pop(i)
                    # 因為我們把當前 i 合併到下一組，所以不必改 i
                    continue
        
            i += 1
        return runs


def chunk_pairs(run):
    pairs=[]
    
    for i in range(0,len(run),2):
        if i+1<len(run):
            pairs.append(run[i:i+2])
    return pairs


# entries = coord_and_angles
# degree_min = 60
# degree_max = 110
def select_corner_pairs(entries, degree_min, degree_max):
    eps = 2
    entries = [e for e in entries if abs(e[-1] - 180) > eps]  #175~185度的都先移除
    
    runs = find_consecutive_angle_runs(entries)
    consec = [r for r in runs if len(r)>1]
    len_consec = sum(len(sub) for sub in consec)

    if len_consec >= 8:
        result=[]
        for r in consec:
            result+=chunk_pairs(r)
    # 沒有連續 -> 單點各自包一組
    elif 0 < len_consec <8:
        result=[]
        for data in runs:
            if len(data) == 2:
                result+=chunk_pairs(data)
            elif len(data) == 1:
                if degree_min <= data[0][-1] <= degree_max:
                    result.append(data)
    else:
        result = [[e] for e in entries if degree_min <= e[-1] <= degree_max]
        # 返回符合角度條件的單個角點列表 (扁平列表)
        
    
    filtered = []
    for group in result:
        # 如果這個 group 裡所有的角度都 ≤ 160，就保留
        if all(e[-1] <= 160 for e in group):
            filtered.append(group)
            
    if len(filtered) > 4:
        # (最小角度, 原始索引)
        min_angles = [
            (min(e[-1] for e in grp), idx)
            for idx, grp in enumerate(filtered)
        ]
        # 取出角度最小的 4 條記錄
        smallest = heapq.nsmallest(4, min_angles, key=lambda x: x[0])
        # 按原始索引排序
        keep_idxs = [idx for _, idx in sorted(smallest, key=lambda x: x[1])]
        # 重建 filtered
        filtered = [filtered[i] for i in keep_idxs]
        
    
    return filtered 

     


#### ── 取得圓弧及倒角的切線交點及角平分線角度 ──────────────────────────────
def get_tangent_intersection(A, B, C, D):
    """
    計算邊 AB 和 CD 的延伸線（假設為直線）的交點 E，並計算角 BEC 的角度。
    Args:
        A (tuple): 點 A 的座標 (x1, y1)
        B (tuple): 點 B 的座標 (x2, y2)
        C (tuple): 點 C 的座標 (x3, y3)
        D (tuple): 點 D 的座標 (x4, y4)
    Returns:
        tuple: (交點 E 的座標 (x, y), 角 BEC 的角度（度）)，如果無交點則返回 (None, None)
    """
    # 直線 AB 的方向向量
    d1x = B[0] - A[0]
    d1y = B[1] - A[1]
    
    # 直線 CD 的方向向量
    d2x = D[0] - C[0]
    d2y = D[1] - C[1]
    
    # 計算分母
    denominator = d1x * d2y - d1y * d2x
    
    # 檢查是否平行
    if abs(denominator) < 1e-12:
        return None, None  # 直線平行，無交點
    
    # 計算 t 和 s
    dx = C[0] - A[0]
    dy = C[1] - A[1]
    t = (dx * d2y - dy * d2x) / denominator
    
    # 計算交點座標 E
    x = A[0] + t * d1x
    y = A[1] + t * d1y
    E = (x, y)
    
    # 計算角 BEC
    # 向量 EB 和 EC
    EB = (B[0] - E[0], B[1] - E[1])
    EC = (C[0] - E[0], C[1] - E[1])
    
    # 計算模長
    mag_EB = math.hypot(EB[0], EB[1])
    mag_EC = math.hypot(EC[0], EC[1])
    
    # 處理退化情況
    if mag_EB < 1e-12 or mag_EC < 1e-12:
        return E, 0.0  # 如果 E 與 B 或 C 重合，返回 0 度
    
    # 計算點積
    dot = EB[0] * EC[0] + EB[1] * EC[1]
    
    # 計算夾角（弧度）
    cos_theta = dot / (mag_EB * mag_EC)
    # 確保 cos_theta 在 [-1, 1] 範圍內，避免浮點誤差
    cos_theta = min(1.0, max(-1.0, cos_theta))
    theta = math.acos(cos_theta)
    
    # 轉換為度數
    angle = math.degrees(theta)
    
    return E, angle



#取得單一個切線的角度跟座標
def get_one_line_intersection(sub_corner):
    """
    sub_corner: [(p_prev1, p_cur1, p_next1, b1, ang1),
                 (p_prev2, p_cur2, p_next2, b2, ang2)]
    回傳：X = 兩條切線在 p_cur1、p_cur2 處的交點
           θ = 這兩條切線方向的夾角(度)
    """
    # 拆出第一條切線的三個點
    p1, p2, p3, b1, ang1 = sub_corner[0]
    # 拆出第二條切線的三個點
    p2, p3, p4, b2, ang2 = sub_corner[1]

    # 求交點及夾角
    X, θ = get_tangent_intersection(p1, p2, p3, p4)

    return p2, X, p3, θ

#計算邊為單一角度的角度值
def get_single_angle(sub_corner):
    """
    計算角 ABC 的角度（以度為單位）。
    Args:
        A (tuple): 點 A 的座標 (x1, y1)
        B (tuple): 點 B 的座標 (x2, y2)
        C (tuple): 點 C 的座標 (x3, y3)
    Returns:
        float: 角 ABC 的角度（度），範圍 [0, 180]
    """
    A = sub_corner[0][0]
    B = sub_corner[0][1]
    C = sub_corner[0][2]
    # 計算向量 BA 和 BC
    BA = (A[0] - B[0], A[1] - B[1])
    BC = (C[0] - B[0], C[1] - B[1])
    
    # 計算模長
    mag_BA = math.hypot(BA[0], BA[1])
    mag_BC = math.hypot(BC[0], BC[1])
    
    # 處理退化情況
    if mag_BA < 1e-12 or mag_BC < 1e-12:
        return 0.0  # 如果 B 與 A 或 C 重合，返回 0
    
    # 計算點積
    dot = BA[0] * BC[0] + BA[1] * BC[1]
    
    # 計算夾角（弧度）
    cos_theta = dot / (mag_BA * mag_BC)
    # 確保 cos_theta 在 [-1, 1] 範圍內，避免浮點誤差
    cos_theta = min(1.0, max(-1.0, cos_theta))
    theta = math.acos(cos_theta)
    
    # 轉換為度數
    angle = math.degrees(theta)
    return angle


#### ── 取得兩兩交點的座標 ──────────────────────────────
#角平分線的方向，就是兩條單位向量的和
def angle_bisector_direction(A, B, C):
    # 1. 向量 BA, BC
    BA = (A[0]-B[0], A[1]-B[1])
    BC = (C[0]-B[0], C[1]-B[1])
    # 2. 單位向量 u1, u2
    def normalize(v):
        mag = math.hypot(v[0], v[1])
        return (v[0]/mag, v[1]/mag) if mag>0 else (0,0)
    u1 = normalize(BA)
    u2 = normalize(BC)
    # 3. 相加
    vx, vy = u1[0]+u2[0], u1[1]+u2[1]
    # 4. 再標準化
    vhat = normalize((vx, vy))
     
    return vhat  # 這就是角平分線的方向單位向量

def intersect_rays(P1, d1, P2, d2, epsilon=1e-9):
    """
    計算兩條射線 P1+t*d1 和 P2+s*d2 的交點。
    P1, P2: 起點 (x,y)
    d1, d2: 單位方向向量 (dx,dy)
    回傳 (x,y) 交點，若平行或無正交點則拋錯。
    """
    x1,y1 = P1
    x2,y2 = P2
    dx1,dy1 = d1
    dx2,dy2 = d2

    # 1. 行列式
    den = dx1*dy2 - dy1*dx2
    if abs(den) < epsilon:
        raise ValueError("兩條射線平行或無交點")

    # 2. 右邊常數
    delta_x = x2 - x1
    delta_y = y2 - y1

    # 3. 求 t1, t2
    t1 = ( delta_x*dy2 - delta_y*dx2 ) / den
    t2 = ( delta_x*dy1 - delta_y*dx1 ) / den

    # 4. 必須在「射線」方向上才算真正交點
    if t1 < 0 or t2 < 0:
        # raise ValueError("交點不在射線的正向上")
        return [], [], []

    # 5. 計算交點座標
    xi = x1 + t1*dx1
    yi = y1 + t1*dy1
    return (xi, yi), t1, t2



def get_intersections_points(intersection_and_angle_dict):    
    # 檢查是否有足夠的角點（至少 4 個）
    n = len(intersection_and_angle_dict)
    if n < 4:
        return []  # 少於 4 個點，無法形成 E 和 F

    # 定義分組：(0, 1) 形成 E，(2, 3) 形成 F
    intersections = []

    for i in range(n):
        prev_key = (i - 1) % n  
        curr_key = i
        next_key = (i + 1) % n  
        
        #射線1座標, 射線交點座標, 射線3座標, 角度
        prev_point_1, prev_point_2, prev_point_3, prev_θ = intersection_and_angle_dict[prev_key]
        curr_point_1, curr_point_2, curr_point_3, curr_θ = intersection_and_angle_dict[curr_key]
        next_point_1, next_point_2, next_point_3, next_θ = intersection_and_angle_dict[next_key]
        
        #射線單位向量
        prev_unit_ray = angle_bisector_direction(prev_point_1, prev_point_2, prev_point_3)
        curr_unit_ray = angle_bisector_direction(curr_point_1, curr_point_2, curr_point_3)
        next_unit_ray = angle_bisector_direction(next_point_1, next_point_2, next_point_3)
        
        
        intersection_prev, t1_curr, t2_prev = intersect_rays(curr_point_2, curr_unit_ray, prev_point_2, prev_unit_ray)
        intersection_next, t2_curr, t2_next = intersect_rays(curr_point_2, curr_unit_ray, next_point_2, next_unit_ray) 
        
        if intersection_prev == [] or intersection_next == []:
            continue
        
        mid_point = [(curr_point_1[0]+curr_point_3[0])/2, (curr_point_1[1]+curr_point_3[1])/2]
        
        if t2_prev > t2_next:
            intersection_data = [(curr_key, next_key), curr_point_2, intersection_next, curr_unit_ray, mid_point, t1_curr, t2_next]
        else:
            intersection_data = [(curr_key, prev_key), curr_point_2, intersection_prev, curr_unit_ray, mid_point, t2_curr, t2_prev]
            
        intersections.append(intersection_data)
       
    return intersections  #[與第幾個點相交的 key, 切線交點, 兩射線在polyline交點, 射線單位向量, 中點
                          #自己切線交點到內部交點的長度, 相交點切線交點到內部交點的長度]


# intersections_list = intersections
#處理配對問題，若已經有(1,2)、(2,1)、「(0,3)、(3,2)」 -> (1,2)、(2,1)、「(0,3)、(3,0)」
def pair_intersections(intersections_list):
    """
    接收 get_intersections_points 的輸出，強制配對交點資料。

    Args:
        intersections_list: get_intersections_points 函數的輸出列表，
                            每個元素格式為 [(curr_key, intersecting_key), curr_point_2, 
                                           intersection_point, curr_unit_ray, mid_point, 
                                           t1_curr, t2_other]。

    Returns:
        一個新的列表，其中包含已強制配對的交點資料。
        每個配對包含兩個元素，代表 (i, j) 和 (j, i) 的資料。
        如果無法配對，將打印警告信息。
    """
    n = len(intersections_list)
    if n < 2:
        return [] # 至少需要兩個點才能配對

    paired_results = []
    paired_indices = set() # 記錄已經配對過的原始索引

    # 建立一個字典，方便透過 key 查找原始列表中的索引
    # 假設 key 是從 0 到 n-1 的整數
    key_to_index = {data[0][0]: i for i, data in enumerate(intersections_list)}
    
    for i in range(n):
        if i in paired_indices:
            continue # 如果這個索引已經處理過，跳過

        current_data = intersections_list[i]
        curr_key, initial_target_key = current_data[0]
        current_intersection_point = current_data[2]
        current_len_from_curr = current_data[5] # 自己切線交點到內部交點的長度
        current_len_from_target = current_data[6] # 相交點切線交點到內部交點的長度

        # 嘗試尋找理論上的配對夥伴 (target_key, curr_key)
        # 它的 curr_key 應該是我們的 initial_target_key
        expected_partner_index = key_to_index.get(initial_target_key)

        if expected_partner_index is not None and expected_partner_index not in paired_indices:
            partner_data = intersections_list[expected_partner_index]
            partner_key, partner_target_key = partner_data[0]

            # 檢查找到的夥伴是否確實是 initial_target_key 開頭
            if partner_key != initial_target_key:
                 # 理論上 key_to_index 應該保證這一點，但加個檢查更安全
                 print(f"警告：索引 {i} ({curr_key}, {initial_target_key}) 尋找夥伴 {initial_target_key} 時，在索引 {expected_partner_index} 找到的卻是鍵 {partner_key} 開頭的資料。跳過配對。")
                 continue

            # 情況一：完美配對 (e.g., (1, 2) 找到了 (2, 1))
            if partner_target_key == curr_key:
                paired_results.append(current_data)
                paired_results.append(partner_data)
                paired_indices.add(i)
                paired_indices.add(expected_partner_index)

            # 情況二：強制配對 (e.g., (0, 3) 找到了 (3, 2)，需要強制改成 (3, 0))
            else :

                # 創建修改後的夥伴資料
                modified_partner_data = [
                    (partner_key, curr_key),            # 修正 key
                    partner_data[1],                    # 保留夥伴自己的切線交點
                    current_intersection_point,         # 使用 current_data 的交點
                    partner_data[3],                    # 保留夥伴自己的射線向量
                    partner_data[4],                    # 保留夥伴自己的中點
                    current_len_from_target,            # 夥伴到交點的距離 = current 到交點的距離 (來自 current_data 的第 7 個元素)
                    current_len_from_curr             # current 到交點的距離 = 夥伴到交點的距離 (來自 current_data 的第 6 個元素)
                ]
                paired_results.append(current_data)
                paired_results.append(modified_partner_data)
                paired_indices.add(i)
                paired_indices.add(expected_partner_index) # 標記原始夥伴索引已使用

        else:
            # 找不到夥伴或夥伴已被配對
            if expected_partner_index is None:
                 print(f"警告：索引 {i} ({curr_key}, {initial_target_key}) 找不到鍵為 {initial_target_key} 的夥伴資料。")
            else:
                 print(f"警告：索引 {i} ({curr_key}, {initial_target_key}) 的預期夥伴索引 {expected_partner_index} 已被配對。")
            # 可以選擇將未配對的 current_data 加入結果，或完全忽略
            # modified_partner_data = [
            #     ('-', '-'),            # 修正 key
            #     current_data[1],                    # 保留夥伴自己的切線交點
            #     intersections_list[0][2],         # 使用 current_data 的交點
            #     current_data[3],                    # 保留夥伴自己的射線向量
            #     current_data[4],                    # 保留夥伴自己的中點
            #     current_len_from_target,            # 夥伴到交點的距離 = current 到交點的距離 (來自 current_data 的第 7 個元素)
            #     current_len_from_curr             # current 到交點的距離 = 夥伴到交點的距離 (來自 current_data 的第 6 個元素)
            # ]
            # paired_results.append(modified_partner_data) # 如果需要保留未配對的項目

    # 檢查是否所有項目都被配對了
    if len(paired_indices) != n:
        print(f"警告：並非所有交點資料都成功配對。預期 {n} 個，實際配對 {len(paired_indices)} 個。")
        unpaired_indices = [idx for idx in range(n) if idx not in paired_indices]
        print(f"未配對的原始索引: {unpaired_indices}")


    return paired_results




#### ── 主程式：繪製中垂線 vs. 角平分線 ──────────────────────────────

def arc_sagitta(start, end, bulge):
    """
    計算圓弧的箭高 (sagitta)，
    輸入：
      start: (x1, y1) 圓弧起點
      end:   (x2, y2) 圓弧終點
      bulge: 圓弧的 bulge 值 = tan(θ/4)
    回傳：
      sagitta: 從弦中點到圓弧的距離（黃色線長度）
    """
    # 1. 計算弦長 c
    dx = end[0] - start[0]
    dy = end[1] - start[1]
    c = math.hypot(dx, dy)

    # 2. 箭高 s = (c/2) * bulge
    s = 0.5 * c * bulge
    return s



def draw_corner_lines(doc, path, corner_runs, intersections, layer_name):
    """
    繪製中垂線與角平分線，同時把每條線的資訊存到一個 list of dicts 裡面並回傳。
   
    """
    ms = doc.ModelSpace
    try:
        layer = doc.Layers.Item(layer_name)
    except:
        layer = doc.Layers.Add(layer_name)
    layer.Color = 3

    def make_point(x,y,z=0):
        return VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, [x,y,z])
    def midpoint(a,b):
        return ((a[0]+b[0])/2, (a[1]+b[1])/2)
    def normalize(v):
        mag = math.hypot(v[0], v[1])
        return (v[0]/mag, v[1]/mag) if mag>1e-12 else (0,0)

    pts    = path['points']
    bulges = path['bulges']
    closed = path['closed']
    m = len(pts) - (1 if closed and pts[0]==pts[-1] else 0)

    # 這就是我們要收集的資料 list
    drawn_lines = []
    
    #繪製內部交點連線(目前限制兩個)    
    second_values = [item[2] for item in intersections]
    unique_second_values = list(set(second_values))    
    if len(unique_second_values) == 2:
        line = ms.AddLine(make_point(*unique_second_values[0]), make_point(*unique_second_values[1]))
        line.Layer = layer_name
        line.Update()
    

    if not intersections or len(intersections) < len(corner_runs):
        # 你可以選擇 raise 一個錯誤，或是直接跳過畫交點的邏輯
        return []

    # 既然檢查完畢，這裡 intersections[key] 一定安全
    drawn_lines = []
  
            
    poly_ent = doc.HandleToObject(path['handle'])
    for key, run in enumerate(corner_runs):
        intersections_data = intersections[key]
        mid = intersections_data[4]       # bisector 起點
        end = intersections_data[2]       # bisector 原本的終點（垂足／平分足）
        ux, uy = intersections_data[3]    # bisector 方向
    
        if len(run) == 2:
            # 中垂線 + sagitta
            # 計算 sagitta 長度 s
            _, p1, _, b1, _ = run[0]
            _, p2, _, b2, _ = run[1]
            b = abs(b1) if abs(b1)>1e-9 else b2
            # 用 arc_sagitta(p1,p2,b) 算出 s
            s = arc_sagitta(p1, p2, b)
    
            # 1. 畫一條偵測線：mid 往前後各延伸 L
            L = s + 100.0  # 延伸長度，自行調整足夠跨出紅線
            x0,y0 = mid
            p0 = (x0 - ux*L, y0 - uy*L)
            p2 = (x0 + ux*L, y0 + uy*L)
            temp = ms.AddLine(make_point(*p0), make_point(*p2))
            # 2. IntersectWith 原 polyline
            pts = temp.IntersectWith(poly_ent, 0)  # acExtendNone=0
            # 可刪除 temp 線避免殘留
            temp.Delete()
    
            # 3. 解析交點陣列
            sag_pt = mid
            if pts:
                arr = list(pts)
                # arr = [x1,y1,z1, x2,y2,z2, ...]
                cands = [(arr[i], arr[i+1]) for i in range(0,len(arr),3)]
                # 取最近 mid 的那個
                cands.sort(key=lambda p: math.hypot(p[0]-x0, p[1]-y0))
                sag_pt = cands[0]
            
            
            
            offset = 0.1
            sag_pt = (sag_pt[0] - ux*offset, sag_pt[1] - uy*offset)
            # 4. 在 CAD 畫真正的 sagitta
            ln = ms.AddLine(make_point(*end), make_point(*sag_pt))
            ln.Layer = layer_name
            ln.Update()
    
            drawn_lines.append({
                "type": "perpendicular_with_sagitta",
                "mid": mid,
                "perp_end": end,
                "sagitta": s,
                "boundary_point": sag_pt,
                "run": run
            })
     

        # —— 單點一組：角平分線 —— 
        else:
            p_prev, p_cur, p_next, bulge, angle = run[0]

            # 找 idx
            i = next((j for j in range(m) if pts[j]==p_cur), None)
            if i is None:
                continue

            # 起點就是 p_cur
            start_pt = p_cur
            end_pt   = intersections_data[2]

            bl = ms.AddLine(make_point(*start_pt), make_point(*end_pt))
            bl.Layer = layer_name
            bl.Update()

            drawn_lines.append({
                "type":  "bisector",
                "start": start_pt,
                "end":   end_pt,
                "angle": angle,
                "boundary_point": p_cur,
                "run":   run
            })

    return drawn_lines




# ── 執行範例 ──────────────────────────────────────────────────
print("🔎 繪製角平分線中...")
layer_name = f'bisector_line'

bisector_dict = {}
a = []
for path in polylines:
    handle = path['handle']
    points = path['points']
    bulges = path['bulges']
    n = len(points)
    coord_and_angles = []
    
    # 從 i=1 到 i=n-2（非閉合時不算首尾；閉合時會自動處理）
    for i in range(len(points)-1):
        if i ==0:
            p_prev = points[i-2]
        else:
            p_prev = points[i-1]
        p_cur  = points[i]
        p_next = points[i+1]

        # 組向量
        v1 = (p_prev[0] - p_cur[0], p_prev[1] - p_cur[1])
        v2 = (p_next[0] - p_cur[0], p_next[1] - p_cur[1])

        # 計算角度
        θ = angle_between(v1, v2)
        coord_and_angles.append([p_prev, p_cur, p_next, bulges[i], round(θ, 2)])
        
        #計算最長距離
        # 計算三邊長
        d1 = hypot(p_prev[0]-p_cur[0], p_prev[1]-p_cur[1])
        d2 = hypot(p_cur[0]-p_next[0], p_cur[1]-p_next[1])
        d3 = hypot(p_prev[0]-p_next[0], p_prev[1]-p_next[1])

    # 2. 分組
    corner_runs = select_corner_pairs(coord_and_angles, degree_min=60, degree_max=110)
    
    #圓弧及倒角的切線交點和角平分線角度
    intersection_and_angle_dict = {}
    #全都是兩兩一組(各個角都是圓弧或倒角)
    if all(len(run) == 2 for run in corner_runs):
        for i in range(len(corner_runs)):
            sub_corner = corner_runs[i]   
            prev_point, intersection_point, next_point, θ = get_one_line_intersection(sub_corner)
            intersection_and_angle_dict[i] = [prev_point, intersection_point, next_point, θ]
    #全都是單一個一組(各個角都是單一角)
    elif all(len(run) == 1 for run in corner_runs):
        for i, sub_corner in enumerate(corner_runs):
            θ = get_single_angle(sub_corner)
            intersection_and_angle_dict[i] = [sub_corner[0][0], sub_corner[0][1], sub_corner[0][2], θ] 
    #包含「圓弧及單一角」或「倒角及單一角」
    else:
        for i, sub_corner in enumerate(corner_runs):
            if len(sub_corner) == 2:
                prev_point, intersection_point, next_point, θ = get_one_line_intersection(sub_corner)
                intersection_and_angle_dict[i] = [prev_point, intersection_point, next_point, θ]
            elif len(sub_corner) == 1:
                θ = get_single_angle(sub_corner)
                intersection_and_angle_dict[i] = [sub_corner[0][0], sub_corner[0][1], sub_corner[0][2], θ] 
 
    #取得角平分線射線的交點
    intersections = get_intersections_points(intersection_and_angle_dict)
    intersections = pair_intersections(intersections)
    a.append(coord_and_angles)
    # 3. 繪製所有角線
    drawn_lines = draw_corner_lines(doc, path, corner_runs, intersections, layer_name)
    bisector_dict[handle] = drawn_lines
    
    
    
boundary_points = [
    info["boundary_point"]
    for lines in bisector_dict.values()
    for info in lines
    if "boundary_point" in info and info["boundary_point"] is not None
]

#%% 繪製道路中心線
from math import atan, sin, cos, pi
from shapely.ops import unary_union
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, LinearRing
from centerline.geometry import Centerline # 匯入 centerline 函式庫
from shapely.ops import linemerge
from shapely.geometry import LineString as ShapelyLine, Point
import win32com.client
from shapely.ops import linemerge, snap

def bulge_to_arc(p1, p2, bulge, segments):
    """
    將一段帶 bulge 的圓弧，近似成多個線段。
    p1, p2: (x,y)
    bulge = tan(theta/4)，theta = sweep angle
    segments: 切分細緻度，bulge 越大可加大
    回傳一系列點（含起點，不含終點）
    """
    if abs(bulge) < 1e-9:
        # 直線段：只回傳起點
        return [p1]

    # 計算弦長與中央角
    dx, dy = p2[0]-p1[0], p2[1]-p1[1]
    chord = (dx*dx + dy*dy)**0.5
    theta = 4 * atan(bulge)  # sweep angle
    radius = chord / (2*sin(theta/2))

    # 圓心
    # 中點
    mx, my = (p1[0]+p2[0])/2, (p1[1]+p2[1])/2
    # 法向量方向
    nx, ny = -dy, dx
    if bulge < 0: 
        nx, ny = -nx, -ny
    # normalize
    d = (nx*nx+ny*ny)**0.5
    nx, ny = nx/d, ny/d
    # h = distance from chord-mid to center
    h = radius * cos(theta/2)
    cx, cy = mx + nx*h, my + ny*h

    # 起訖角度
    import math
    ang1 = math.atan2(p1[1]-cy, p1[0]-cx)
    ang2 = ang1 + theta

    pts = []
    for i in range(segments):
        t = ang1 + (theta * i/segments)
        pts.append((cx + radius*cos(t), cy + radius*sin(t)))
    return pts

def polyline_to_polygon(points, bulges, closed, arc_segments):
    """
    points: list of (x,y)
    bulges: list of float, 與 segments 對應 (最後一段可是 0)
    closed: bool
    arc_segments: 每條 bulge 弧分段數
    回傳 Shapely Polygon
    """
    ring_pts = []
    n = len(points)
    last = n-1 if closed else n-2
    for i in range(n-1 if not closed else n):
        p1 = points[i]
        p2 = points[(i+1)%n]
        b  = bulges[i]
        arc = bulge_to_arc(p1, p2, b, segments=arc_segments)
        ring_pts.extend(arc)
    # 保證最後回到起點
    if ring_pts[0] != ring_pts[-1]:
        ring_pts.append(ring_pts[0])
    # 建 LinearRing 再轉 Polygon
    lr = LinearRing(ring_pts)
    return Polygon(lr)


polys = []
for pl in polylines:
    poly = polyline_to_polygon(pl['points'], pl['bulges'], pl['closed'], arc_segments=1024)
    polys.append(poly)

# 2. 合併所有塊狀區域
valid_polys = []
invalid_handles = []
for i, poly in enumerate(polys): # 假設 polys 是 Polygon 物件列表
    # 或者如果你有 handle 資訊: for handle, poly in red_polygons.items():
    if poly is None:
         print(f"警告: 在索引 {i} 處發現 None 值。") # 或 handle
         continue

    if poly.is_valid:
        valid_polys.append(poly)
    else:
        print(f"警告: 在索引 {i} 處的多邊形無效。嘗試修復...") # 或 handle
        # 嘗試使用 buffer(0) 修復
        fixed_poly = poly.buffer(0)
        if fixed_poly.is_valid and isinstance(fixed_poly, Polygon):
             print(f"   成功修復索引 {i} 的多邊形。")
             valid_polys.append(fixed_poly)
        else:
             print(f"   ❌ 無法修復索引 {i} 的多邊形。類型: {fixed_poly.geom_type}")
             # 記錄無效多邊形的 handle (如果有的話)
             # invalid_handles.append(handle)

# 然後只對有效的多邊形進行 union
if valid_polys:
    street_region = unary_union(valid_polys)
    print(f"成功對 {len(valid_polys)} 個有效多邊形執行 unary_union。")
else:
    print("❌ 沒有有效的多邊形可供合併。")
    street_region = None # 或者一個空的 GeometryCollection

outer = street_region.convex_hull

roads = outer.difference(street_region)

# 計算中心線，interpolation_distance 控制輸出線條的平滑度/點密度
# 值越小，點越密，線條越平滑，但計算量越大
center_line = Centerline(roads, interpolation_distance=1) # 直接傳入 Shapely 物件

road_skeleton = center_line.geometry



def draw_skeleton_as_polylines(
    doc,
    skeleton,
    layer_name,
    min_length=30,
    tolerance=0.1  # snap 容差
):
    """
    在 CAD 上画出骨架：
     - 先 snap 粘合端点，再 linemerge 生成连续线
     - 过滤长度小于 min_length 的段
    """
    pythoncom.CoInitialize()
    ms = doc.ModelSpace

    # 取或建图层
    try:
        lyr = doc.Layers.Item(layer_name)
    except:
        lyr = doc.Layers.Add(layer_name)
    lyr.Color = 6

    # 1. 将 skeleton 标准化为 MultiLineString
    base = skeleton
    if isinstance(base, LineString):
        base = MultiLineString([base])

    # 2. snap：把相距 <= tolerance 的端点“粘合”在一起
    snapped = snap(base, base, tolerance)

    # 3. linemerge：合并成真正的连续段
    merged = linemerge(snapped)

    # 4. 拆成 list of LineString
    if isinstance(merged, LineString):
        lines = [merged]
    elif isinstance(merged, MultiLineString):
        lines = list(merged.geoms)
    else:
        raise TypeError(f"不支持的 geometry: {merged.geom_type}")

    def make_array(coords):
        arr = []
        for x, y in coords:
            arr.extend([x, y, 0.0])
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, arr)

    endpoints = []

    # 5. 逐条处理：过滤、绘制
    for ls in lines:
        if ls.length < min_length:
            continue
        pts = list(ls.coords)

        # 如果你还想检查“近直线”分支，可按之前逻辑再分线/弧
        # 这里简化都用 AddPolyline
        vt = make_array(pts)
        pl = ms.AddPolyline(vt)
        pl.Closed = False
        pl.Layer = layer_name
        pl.Update()

        endpoints.append((pts[0], pts[-1]))

    return endpoints



# ===== 使用範例 =====
endpoints = draw_skeleton_as_polylines(
    doc,
    road_skeleton,
    layer_name="road_central_line",
    min_length=30 #只取大於30m的線
)





#%%繪製街廓邊緣到交叉路口的線


def connect_boundary_to_endpoints(
    doc,
    boundary_points,   # [(x,y), ...]
    endpoints,         # [(start_pt, end_pt), ...]
    layer_name,
    max_dist=20        # 最高連線距離（公尺）
):
    """
    在 CAD 中，將每個 boundary_point 連到 endpoints 中最近的那個點。
    若最短距離超過 max_dist 則跳過不連。
    """
    ms = doc.ModelSpace
    try:
        layer = doc.Layers.Item(layer_name)
    except:
        layer = doc.Layers.Add(layer_name)

    # 將 endpoints 平展成一維點陣列
    flat_eps = []
    for s, e in endpoints:
        flat_eps.append(s)
        flat_eps.append(e)
    eps_array = np.array(flat_eps)  # shape = (N,2)

    tree = KDTree(eps_array)

    def make_pt(pt):
        x, y = pt
        return VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, [x, y, 0.0])

    count = 0
    for bp in boundary_points:
        dist, idx = tree.query(bp)
        if dist > max_dist:
            # 跳過超過 20m 的連線
            continue

        nearest = tuple(eps_array[idx])
        va = make_pt(bp)
        vb = make_pt(nearest)
        ln = ms.AddLine(va, vb)
        ln.Layer = layer_name
        ln.Update()
        count += 1

    print(f"✅ 已在圖層「{layer_name}」連出了 {count} 條距離 ≤ {max_dist}m 的線段。")

# ===== 使用範例 =====
connect_boundary_to_endpoints(doc, boundary_points, endpoints, layer_name = 'bisector_line', max_dist=20)




#%% 提取 bisector_line 圖層的線段


print("🔎 從 bisector_line 圖層提取線段並傳給 unary_union...")

from shapely.geometry import LineString, MultiLineString
from shapely.ops import unary_union, split, polygonize  # 明確導入 split

def extract_lines_from_layer(doc, layer_name):
    model_space = doc.ModelSpace
    lines = []
    for ent in model_space:
        if ent.Layer != layer_name:
            continue
        try:
            if ent.ObjectName == 'AcDbLine':
                sp = ent.StartPoint; ep = ent.EndPoint
                if len(sp) < 2 or len(ep) < 2:
                    continue
                line = LineString([tuple(sp[:2]), tuple(ep[:2])])
                if line.is_valid and line.length > 0:
                    lines.append(line)

            elif ent.ObjectName in ['AcDbPolyline','AcDb2dPolyline']:
                # (1) 頂點數檢查
                try:
                    cnt = ent.NumberOfVertices
                except:
                    cnt = None
                coords = list(ent.Coordinates)
                if (cnt is not None and cnt < 2) or len(coords) < 4:
                    continue

                # (2) 把 coords 轉 (x,y) 點
                pts = [(coords[i],coords[i+1]) for i in range(0,len(coords),2)]

                # (3) 讀 bulge list，保底為 0
                bulges = []
                for i in range(len(pts)-1):
                    try:    b = ent.GetBulge(i)
                    except: b = 0
                    bulges.append(b)
                bulges.append(0)

                # (4) bulge_to_arc：只給合法段
                line_pts = []
                for i,(p1,p2) in enumerate(zip(pts, pts[1:])):
                    bp = bulge_to_arc(p1,p2,bulges[i],segments=16)
                    if len(bp) >= 2:
                        line_pts.extend(bp[:-1])
                line_pts.append(pts[-1])

                line = LineString(line_pts)
                if line.is_valid and line.length > 0:
                    lines.append(line)

        except Exception as e:
            print(f"⚠️ 無法處理物件 Handle={ent.Handle}：{e}")
            continue
    return lines


# 1. 提取 bisector_line 圖層的線段
bisector_layer_name = 'bisector_line'
bisector_lines = extract_lines_from_layer(doc, bisector_layer_name)


#%%取得子集水區面積
import pythoncom
import win32com.client
from win32com.client import VARIANT
from shapely.ops import unary_union, polygonize
from shapely.geometry import MultiLineString, LineString

def annotate_areas(
    doc,
    region,
    bisector_lines,
    skeleton_lines,
    inset_eps,
    layer_name,
    label_prefix,
    min_area,
    text_height
):
    """
    通用的面积标注函数。
    
    region: Shapely Polygon/MultiPolygon，待切割区域
    bisector_lines: List[LineString]，用于切割的 bisector 线
    skeleton_lines: List[LineString]，可选的额外骨架线（道路中心线）
    inset_eps: float，region.buffer(-inset_eps) 的缩进距离
    layer_name: str，CAD 图层名
    label_prefix: str，标注文字前缀，e.g. "住宅"/"道路"
    min_area: float，低于此面积不标注
    text_height: float，CAD 文字高度
    """
    # 初始化 CAD
    pythoncom.CoInitialize()
    acad = win32com.client.Dispatch("AutoCAD.Application")
    ms   = doc.ModelSpace

    # 1. 准备切割线集合
    # 1.1 bisector 切割到 region 内
    cutter = unary_union(bisector_lines).intersection(region)
    cut_segs = []
    if cutter.geom_type == 'MultiLineString':
        cut_segs.extend(cutter.geoms)
    elif cutter.geom_type == 'LineString':
        cut_segs.append(cutter)

    # 1.2 加入骨架线
    for sk in skeleton_lines:
        if isinstance(sk, (LineString,)):
            cut_segs.append(sk)
        elif isinstance(sk, MultiLineString):
            cut_segs.extend(sk.geoms)

    # 1.3 取 region 缩进后边界
    inset = region.buffer(-inset_eps)
    bnd = inset.boundary
    if bnd.geom_type == 'MultiLineString':
        cut_segs.extend(bnd.geoms)
    else:
        cut_segs.append(bnd)

    # 2. polygonize
    net = unary_union(cut_segs)
    all_pieces = list(polygonize(net))

    # 3. 裁剪到原始 region
    subregions = [
        p.intersection(region)
        for p in all_pieces
        if p.intersects(region)
    ]

    # 4. 建立／切换到图层
    layers = doc.Layers
    try:
        layers.Item(layer_name)
    except:
        layers.Add(layer_name)

    # 5. 标注文字
    for poly in subregions:
        a = poly.area
        if a <= min_area:
            continue
        pt = poly.centroid  # 也可用 representative_point()
        x, y = pt.x, pt.y
        txt = f"{label_prefix} area:{a:.4f} m2"
        ins = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, 0.0))
        ent = ms.AddText(txt, ins, text_height)
        ent.Update()

    # 6. 刷新视图
    acad.ActiveDocument.Regen(0)
    print(f"✅ 已在 CAD 圖層「{layer_name}」標註 {label_prefix} 面積")

    # 返回 subregions 以备后续分析
    return subregions


# ===== 使用示例 =====

# （1）住宅区面积
street_subs = annotate_areas(
    doc=doc,
    region=street_region,
    bisector_lines=bisector_lines,
    skeleton_lines=[],         
    inset_eps=1e-9,
    layer_name="street_area",
    label_prefix="住宅",
    min_area=0,
    text_height=0.8
)

# （2）道路区面积
# 这里先准备切割线：bisector ∩ roads，再骨架线 road_skeleton
cut_road = unary_union(bisector_lines).intersection(roads)
road_cut_segs = []
if cut_road.geom_type=='MultiLineString':
    road_cut_segs.extend(cut_road.geoms)
elif cut_road.geom_type=='LineString':
    road_cut_segs.append(cut_road)


road_subs = annotate_areas(
    doc=doc,
    region=roads,
    bisector_lines=road_cut_segs,
    skeleton_lines=[road_skeleton],
    inset_eps=1e-8,
    layer_name="road_area",
    label_prefix="道路",
    min_area=100,
    text_height=0.8
)

 


#%% 繪製道路側溝( 輸入「平移距離」及「側溝寬度」)
import math
import pythoncom
import win32com.client
from win32com.client import VARIANT


def vertex_angle(p_prev, p_cur, p_next):
    """
    計算三點在 p_cur 處所構成的內角，回傳度數 (0~180)。
    
    參數：
      - p_prev: (x, y)   前一個頂點
      - p_cur : (x, y)   中間頂點
      - p_next: (x, y)   後一個頂點
    
    回傳：
      - 角度 (float)，若任一線段過短則回傳 None
    """
    # 向量 v1: p_cur -> p_prev，v2: p_cur -> p_next
    v1x = p_prev[0] - p_cur[0]
    v1y = p_prev[1] - p_cur[1]
    v2x = p_next[0] - p_cur[0]
    v2y = p_next[1] - p_cur[1]
    
    # 長度
    L1 = math.hypot(v1x, v1y)
    L2 = math.hypot(v2x, v2y)
    if L1 < 1e-9 or L2 < 1e-9:
        return None  # 線段太短，無法計算
    
    # 內積與 cosθ
    dot = v1x * v2x + v1y * v2y
    cos_theta = dot / (L1 * L2)
    # 避免浮點誤差超出 [-1,1]
    cos_theta = max(-1.0, min(1.0, cos_theta))
    
    # 角度 (rad) → 角度 (deg)
    theta_rad = math.acos(cos_theta)
    theta_deg = math.degrees(theta_rad)
    return theta_deg





def batch_offset_polylines(
    doc,
    polylines_info,
    offset_dist,
    width,
    square_width,
    draw_junction_angle,
    insetsize,
    dst_layer='test_off'
):
    """
    批次對多段線做 Offset，並在偏移距離 = offset_dist 的中心線上，
    每個頂點繪製旋轉方塊。若該頂點對應的 bulge ≠ 0，則用「前一段向量」
    而非「下一段向量」來計算旋轉角度。

    參數說明同上，此處省略。
    """
    ms = doc.ModelSpace

    #── 確保目標圖層存在且解鎖 ──#
    try:
        lyr = doc.Layers.Item(dst_layer)
        lyr.Lock = False
    except:
        lyr = doc.Layers.Add(dst_layer)

    #── 確保虛線線型存在 ──#
    dashed = "DASHED"
    try:
        doc.Linetypes.Item(dashed)
    except:
        try:    doc.Linetypes.Load(dashed, "acad.lin")
        except: dashed = "CONTINUOUS"

    success_count = 0
    square_count  = 0

    #── 逐條多段線處理 ──#
    for info in polylines_info:
        h = info['handle']
        try:
            ent = doc.HandleToObject(h)
        except:
            print(f"⚠️ 找不到 handle={h}，跳過")
            continue
        if ent.ObjectName not in ('AcDbPolyline','AcDbLine'):
            continue

        #── 計算欲偏移的距離串 ──#
        if width > 0:
            offsets = [offset_dist+width/2,
                       offset_dist-width/2,
                       offset_dist]
        else:
            offsets = [offset_dist]

        main_offset_entities = []

        #── 執行 OFFSET──#
        for dist in offsets:
            try:
                res = ent.Offset(dist)
            except:
                continue
            ents = list(res) if isinstance(res,(tuple,list)) else [res]
            for ne in ents:
                ne.Layer = dst_layer
                if abs(dist-offset_dist) < 1e-6:
                    ne.Linetype      = dashed
                    ne.LinetypeScale = 1.0
                    main_offset_entities.append(ne)
                ne.Update()
                success_count += 1

        #── 畫集水井 ──#
        if square_width > 0:
            half = square_width / 2

            for ne in main_offset_entities:
                if not hasattr(ne, "Coordinates"):
                    continue
                # 讀出所有頂點
                arr  = list(ne.Coordinates)
                pts2 = [(arr[i*2], arr[i*2+1]) for i in range(len(arr)//2)]

                # 一一對每個頂點畫方塊
                for j, (cx, cy) in enumerate(pts2):
                    # 取 bulge（針對 polyline）
                    bulge = 0
                    if ent.ObjectName == 'AcDbPolyline':
                        try:
                            bulge = ne.GetBulge(j)
                        except:
                            bulge = 0
                                                       
                    p_prev = pts2[j-1] 
                    p_cur  = pts2[j]
                    p_next = pts2[j+1] if j<len(pts2)-1 else pts2[0]
                    angle = vertex_angle(p_prev, p_cur, p_next)
                    
                    #設置轉折點在90~160度才畫集水井
                    if draw_junction_angle[0] <= angle <= draw_junction_angle[1]:
                        # 計算「前一段」與「下一段」向量
                        if bulge == 0:
                            nx, ny = pts2[j+1] if j<len(pts2)-1 else pts2[0]
                            dx, dy = nx - cx, ny - cy
                        elif bulge!= 0:
                            px, py = pts2[j-1]
                            dx, dy = cx - px, cy - py                      
    
                        # 將向量標準化
                        L = math.hypot(dx, dy)
                        if L < 1e-6:
                            dx, dy = 1, 0
                        else:
                            dx /= L; dy /= L
    
                        # 計算與水平線的弧度夾角
                        angle = math.atan2(dy, dx)
                        # （bulge 情況已由上面邏輯決定反向向量）
    
                        # # 根據此角度把本地正方形頂點旋轉到世界座標
                        # va = []
                        # ct, st = math.cos(angle), math.sin(angle)
                        # for lx, ly in [(-half,-half),( half,-half),
                        #                ( half, half),(-half, half)]:
                        #     xw = cx + lx*ct - ly*st
                        #     yw = cy + lx*st + ly*ct
                        #     va.extend([xw, yw])
    
                        # # 畫出封閉多段線（正方形）
                        # arr_va = VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, va)
                        # sq = ms.AddLightWeightPolyline(arr_va)
                        # sq.Closed = True
                        # sq.Layer  = dst_layer
                        # sq.Update()
                        # square_count += 1
                        # 計算好旋轉角度 θ 之後，先畫「外層」正方形
                        va = []
                        ct, st = math.cos(angle), math.sin(angle)
                        for lx, ly in [(-half,-half),( half,-half),
                                       ( half, half),(-half, half)]:
                            xw = cx + lx*ct - ly*st
                            yw = cy + lx*st + ly*ct
                            va.extend([xw, yw])
        
                        arr_va = VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, va)
                        sq = ms.AddLightWeightPolyline(arr_va)
                        sq.Closed = True
                        sq.Layer  = dst_layer
                        sq.Update()
                        square_count += 1
        
                        #── 接著畫「內層」同心正方形──#
                        inner_half = half - insetsize
                        if inner_half > 0:
                            va2 = []
                            for lx, ly in [(-inner_half,-inner_half),
                                           ( inner_half,-inner_half),
                                           ( inner_half, inner_half),
                                           (-inner_half, inner_half)]:
                                xw = cx + lx*ct - ly*st
                                yw = cy + lx*st + ly*ct
                                va2.extend([xw, yw])
        
                            arr_va2 = VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, va2)
                            sq2 = ms.AddLightWeightPolyline(arr_va2)
                            sq2.Closed = True
                            sq2.Layer  = dst_layer
                            sq2.Color  = sq.Color  # 跟外層一樣顏色
                            sq2.Update()
        
                            #── 在內層正方形畫 X──#
                            # 頂點順序：0:(-,-), 1:(+,-), 2:(+,+), 3:(-,+)
                            # X 由 0→2 與 1→3 兩條線組成
                            p0 = (va2[0], va2[1])
                            p1 = (va2[2], va2[3])
                            p2 = (va2[4], va2[5])
                            p3 = (va2[6], va2[7])
        
                            ln1 = ms.AddLine(VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, (*p0,0.0)),
                                             VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, (*p2,0.0)))
                            ln2 = ms.AddLine(VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, (*p1,0.0)),
                                             VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, (*p3,0.0)))
                            for ln in (ln1, ln2):
                                ln.Layer = dst_layer
                                ln.Color = sq.Color
                                ln.Update()
        
                            square_count += 1  # 也可不算入次數

    #── 完成後刷新視窗 ──#
    doc.Regen(0)


# 示例调用
batch_offset_polylines(
    doc,
    polylines_info=polylines,
    offset_dist=1.0,
    width=1,      # 側溝寬度
    square_width=1.2,  # 集水井邊長
    draw_junction_angle = [90, 160],
    insetsize=0.2,
    dst_layer='test_off'
)