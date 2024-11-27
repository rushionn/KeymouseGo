# 在文件開頭加入必要的 imports
from datetime import datetime
import time
import logging
from typing import List, Tuple

class EnhancedRecorder:
    def __init__(self):
        self.recording_data = []
        self.is_recording = False
        self.mouse_precision = 2  # 增加滑鼠精確度設定
        
    def start_recording(self):
        self.is_recording = True
        self.recording_start_time = time.time()
        
    def record_mouse_movement(self, x, y):
        # 優化滑鼠軌跡記錄，增加平滑度
        if self.is_recording:
            current_time = time.time()
            if len(self.recording_data) > 0:
                last_pos = self.recording_data[-1]
                if self.calculate_distance(last_pos, (x, y)) > self.mouse_precision:
                    self.recording_data.append((x, y, current_time))
