import pandas as pd
import binascii
import os
import math

class UniversalJT808Decoder:
    def __init__(self, config_file='master_mapping_config.csv'):
        self.config = None
        self.rules_map = {} 
        self.all_field_names = [] 
        
        self.status_definitions = {
            0: ("ACC:OFF", "ACC:ON"), 1: ("GPS:NoFix", "GPS:Fixed"),
            2: ("Lat:North", "Lat:South"), 3: ("Lng:East", "Lng:West"),
            10: ("Oil:Normal", "Oil:Disconnect"), 4: ("Op:Normal", "Op:Stop"), 5: ("Enc:No", "Enc:Yes")
        }
        self.alarm_definitions = {
            0: "SOS", 1: "OverSpeed", 7: "MainPower:LowVolt", 8: "MainPower:Off", 29: "Collision",
            2: "Fatigue", 3: "EarlyWarn", 4: "GNSS:Fault", 5: "GNSS:Cut", 6: "Power:Low",
            9: "LCD:Cut", 18: "Drive:Time", 19: "Stop:Time", 20: "Area:In/Out"
        }

        if os.path.exists(config_file):
            try:
                # [FIX 1] อ่าน CSV เป็น String ทั้งหมด เพื่อรักษาเลข 0 นำหน้า
                self.config = pd.read_csv(config_file, dtype=str)
                self.config.fillna('', inplace=True)
                
                # [FIX 2] Smart Padding Function
                def smart_pad(val):
                    val = str(val).strip()
                    # ถ้าเป็น Hex (เช่น 2D, FB) หรือตัวเลข
                    is_hex = all(c in '0123456789ABCDEFabcdef' for c in val)
                    if not is_hex: return val

                    # ถ้ามีความยาวเกิน 2 หรือค่าเกิน 255 (FF) ให้ถือว่าเป็น 2 Bytes -> Pad 4
                    # กรณี BSJ 0110 (Saved as 110) -> Len 3 -> Pad 4 -> 0110 OK
                    # กรณี Ext 01 (Saved as 1) -> Len 1 -> Pad 2 -> 01 OK
                    try:
                        int_val = int(val, 16)
                        if int_val > 255 or len(val) > 2:
                            return val.zfill(4)
                        return val.zfill(2)
                    except:
                        return val

                self.config['SubID'] = self.config['SubID'].apply(smart_pad)
                
                # แปลงค่าตัวเลขที่จำเป็น
                # ต้องทำ manual เพราะเราอ่านมาเป็น str ทั้งหมด
                self.config['StartByte'] = pd.to_numeric(self.config['StartByte'], errors='coerce').fillna(0).astype(int)
                self.config['Length'] = pd.to_numeric(self.config['Length'], errors='coerce').fillna(0).astype(int)
                
                for _, row in self.config.iterrows():
                    cat = row['Category']
                    sub = row['SubID']
                    
                    if cat == 'Standard': key = 'Standard_Rules'
                    elif cat == 'Header': key = f"Header_{sub}" 
                    else: key = f"{cat}{sub.upper()}"
                    
                    if key not in self.rules_map: self.rules_map[key] = []
                    self.rules_map[key].append(row.to_dict())

                self.all_field_names = self.config['FieldName'].tolist()
                
                if '_Status' not in self.all_field_names:
                    self.all_field_names.append('_Status')
                if 'Raw Hex Block' not in self.all_field_names:
                    self.all_field_names.append('Raw Hex Block')
                        
            except Exception as e: print(f"Error reading CSV: {e}")

    def get_header_name(self, sub_id, default_name):
        key = f"Header_{sub_id}"
        if key in self.rules_map:
            return self.rules_map[key][0]['FieldName']
        return default_name

    def unescape(self, content_hex):
        return content_hex.replace('7D02', '7E').replace('7D01', '7D')

    def verify_checksum(self, data_bytes):
        try:
            if len(data_bytes) < 2: return False
            calc = 0
            for b in data_bytes[:-1]: calc ^= b
            return (calc == data_bytes[-1])
        except: return False

    def format_timestamp(self, bcd):
        if not bcd or len(bcd) != 12: return bcd
        return f"20{bcd[0:2]}-{bcd[2:4]}-{bcd[4:6]} {bcd[6:8]}:{bcd[8:10]}:{bcd[10:12]}"

    def parse_value(self, val_hex, rule):
        try:
            dtype = rule['DataType']; scale = float(rule['Scale']) if rule['Scale'] else 1.0
            
            if dtype == 'BCD': return self.format_timestamp(val_hex)
            elif dtype in ['WORD', 'DWORD', 'BYTE']: return int(val_hex, 16) / scale
            elif dtype == 'SIGNED_WORD':
                raw = int(val_hex, 16); return (raw - 0x10000 if raw >= 0x8000 else raw) / scale
            elif dtype == 'SIGNED_DWORD':
                raw = int(val_hex, 16); return (raw - 0x100000000 if raw >= 0x80000000 else raw) / scale
            elif dtype == 'STRING':
                return binascii.unhexlify(val_hex).decode('gbk', errors='ignore').strip('\x00').strip()
            elif dtype == 'HEX': return val_hex
            return val_hex
        except: return f"ERR:{val_hex}"

    def parse_bsj_tlv(self, container_hex):
        res = {}; idx = 0; limit = len(container_hex)
        while idx < limit:
            try:
                if idx+4 > limit: break
                length_val = int(container_hex[idx:idx+4], 16)
                if idx+8 > limit: break 
                sub_id = container_hex[idx+4:idx+8]
                data_len_hex = (length_val - 2) * 2
                data_start = idx+8
                if data_start + data_len_hex > limit: break
                val_hex = container_hex[data_start : data_start + data_len_hex]
                
                key = f"BSJ_Ext{sub_id.upper()}"
                if key in self.rules_map:
                    for rule in self.rules_map[key]:
                        try:
                            s_byte = int(rule.get('StartByte', 0)); l_byte = int(rule.get('Length', 0))
                            if l_byte == 0: l_byte = len(val_hex) // 2
                            s_idx = s_byte * 2; e_idx = s_idx + (l_byte * 2)
                            if e_idx > len(val_hex): e_idx = len(val_hex)
                            if s_idx < len(val_hex):
                                sub_val = val_hex[s_idx : e_idx]
                                res[rule['FieldName']] = self.parse_value(sub_val, rule)
                        except: pass
                idx += 4 + (length_val * 2) 
            except: break
        return res

    def decode_raw(self, raw_hex):
        try:
            error_msgs = []
            hex_str = str(raw_hex).upper().replace(" ", "").replace("\n", "").replace("\r", "").strip()
            if not (hex_str.startswith('7E') and hex_str.endswith('7E')): error_msgs.append("No 7E Header/Tail")

            try:
                content_hex = hex_str[2:-2] if (hex_str.startswith('7E') and hex_str.endswith('7E')) else hex_str
                data_str = self.unescape(content_hex)
                data_bytes = bytes.fromhex(data_str)
            except:
                return [{'Message ID': 'HEX ERR', '_Status': "Invalid Hex", 'Raw Hex Block': hex_str}], "Error", 0

            if not self.verify_checksum(data_bytes): error_msgs.append("Checksum Fail")
            if error_msgs: return [{'Message ID': f"0x{data_str[0:4]}" if len(data_str)>=4 else '?', '_Status': "❌ " + " | ".join(error_msgs), 'Raw Hex Block': hex_str}], "Error", 0

            row_status = "OK"
            data = data_str
            body_attr = int(data[4:8], 16)
            is_2019 = (body_attr >> 14) & 1
            
            header = {}
            header[self.get_header_name('MsgID', 'Message ID')] = f"0x{data[0:4]}"
            header[self.get_header_name('BodyAttr', 'Body Attr Raw')] = f"0x{body_attr:04X}"
            header[self.get_header_name('BodyLen', 'Body Length')] = body_attr & 0x03FF
            
            if is_2019:
                header[self.get_header_name('ProtoVer', 'Protocol Version')] = int(data[8:10], 16)
                header[self.get_header_name('TermID', 'CCID')] = data[10:30]
                header[self.get_header_name('TrackCount', 'Count Up Track')] = int(data[30:34], 16)
                body_start = 34
            else:
                header[self.get_header_name('ProtoVer', 'Protocol Version')] = "-"
                header[self.get_header_name('TermID', 'CCID')] = data[8:20]
                header[self.get_header_name('TrackCount', 'Count Up Track')] = int(data[20:24], 16)
                body_start = 24
            if (body_attr >> 13) & 1: body_start += 8

            def parse_one_msg(b_hex):
                res = {}
                std_len_hex = 56 
                standard_hex = b_hex[:std_len_hex]
                ext_hex = b_hex[std_len_hex:]
                
                if 'Standard_Rules' in self.rules_map:
                    for rule in self.rules_map['Standard_Rules']:
                        try:
                            s_byte = int(rule.get('StartByte', 0)); l_byte = int(rule.get('Length', 0))
                            s_idx = s_byte * 2; e_idx = s_idx + (l_byte * 2)
                            if e_idx <= len(standard_hex):
                                sub_val = standard_hex[s_idx : e_idx]
                                res[rule['FieldName']] = self.parse_value(sub_val, rule)
                        except: pass
                
                i = 0
                while i < len(ext_hex):
                    try:
                        if i+4 > len(ext_hex): break
                        eid = ext_hex[i:i+2]; e_len = int(ext_hex[i+2:i+4], 16)
                        if i+4+(e_len*2) > len(ext_hex): break
                        e_val = ext_hex[i+4 : i+4+(e_len*2)]
                        
                        if eid in ['EB', 'EC', 'ED', 'EF']: res.update(self.parse_bsj_tlv(e_val))
                        else:
                            key = f"Extension{eid.upper()}"
                            if key in self.rules_map:
                                for rule in self.rules_map[key]:
                                    try:
                                        s_byte = int(rule.get('StartByte', 0)); l_byte = int(rule.get('Length', 0))
                                        if l_byte == 0: l_byte = len(e_val) // 2
                                        s_idx = s_byte * 2; e_idx = s_idx + (l_byte * 2)
                                        if e_idx > len(e_val): e_idx = len(e_val)
                                        if s_idx < len(e_val):
                                            sub_val = e_val[s_idx : e_idx]
                                            res[rule['FieldName']] = self.parse_value(sub_val, rule)
                                    except: pass
                        i += 4 + (e_len * 2)
                    except: break
                return {**header, **res, 'Raw Hex Block': b_hex}

            decoded_list = []
            mid = int(data[0:4], 16); body_hex = data[body_start:-2]
            if mid == 0x0704:
                try:
                    count = int(body_hex[0:4], 16); curr = 6
                    for _ in range(count):
                        l_len = int(body_hex[curr:curr+4], 16)
                        decoded_list.append(parse_one_msg(body_hex[curr+4 : curr+4+(l_len*2)]))
                        curr += 4 + (l_len*2)
                except: decoded_list.append(parse_one_msg(body_hex))
            else: decoded_list.append(parse_one_msg(body_hex))
            for item in decoded_list: item['_Status'] = row_status
            return decoded_list, row_status, mid
        except Exception as e: return [{'Message ID': 'CRASH', '_Status': f"❌ Sys Error: {str(e)}"}], "Error", 0