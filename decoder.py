import pandas as pd
import binascii
import os
import math
import re

class UniversalJT808Decoder:
    def __init__(self, config_file='master_mapping_config.csv'):
        self.config = None
        self.rules_map = {} 
        self.all_field_names = [] 
        self.calculated_rules = [] 
        self.container_ids = [] 

        if os.path.exists(config_file):
            try:
                self.config = pd.read_csv(config_file, dtype=str)
                self.config.fillna('', inplace=True)
                
                def smart_pad(val):
                    val = str(val).strip()
                    is_hex = all(c in '0123456789ABCDEFabcdef' for c in val)
                    if not is_hex: return val
                    try:
                        int_val = int(val, 16)
                        if int_val > 255 or len(val) > 2: return val.zfill(4)
                        return val.zfill(2)
                    except: return val

                self.config['SubID'] = self.config['SubID'].apply(smart_pad)
                self.config['StartByte'] = pd.to_numeric(self.config['StartByte'], errors='coerce').fillna(0).astype(int)
                self.config['Length'] = pd.to_numeric(self.config['Length'], errors='coerce').fillna(0).astype(int)
                
                valid_fields = [] 
                
                for _, row in self.config.iterrows():
                    cat = str(row['Category']).strip()
                    sub = str(row['SubID']).strip()
                    
                    if cat.upper() == 'CONTAINER':
                        self.container_ids.append(sub.upper())
                        continue
                    
                    valid_fields.append(row['FieldName'])
                        
                    if row['DataType'] == 'BIT':
                        self.calculated_rules.append(row.to_dict())
                        continue
                    
                    if cat == 'Standard': key = 'Standard_Rules'
                    elif cat == 'Header': key = f"Header_{sub}"
                    elif cat in ['Extension', 'Main']: key = f"Main_{sub.upper()}"
                    elif str(cat).startswith('Sub') or cat == 'BSJ_Ext': 
                        key = f"Sub_{sub.upper()}"      
                    else: key = f"{cat}{sub.upper()}"
                    
                    if key not in self.rules_map: self.rules_map[key] = []
                    self.rules_map[key].append(row.to_dict())

                self.all_field_names = valid_fields
                if '_Status' not in self.all_field_names: self.all_field_names.append('_Status')
                if 'Raw Hex Block' not in self.all_field_names: self.all_field_names.append('Raw Hex Block')
                        
            except Exception as e: print(f"Error reading CSV: {e}")

    def get_header_name(self, sub_id, default_name):
        key = f"Header_{sub_id}"
        if key in self.rules_map: return self.rules_map[key][0]['FieldName']
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
            elif dtype in ['WORD', 'DWORD', 'BYTE', 'INT']: return int(val_hex, 16) / scale
            elif dtype == 'SIGNED_WORD':
                raw = int(val_hex, 16); return (raw - 0x10000 if raw >= 0x8000 else raw) / scale
            elif dtype == 'SIGNED_DWORD':
                raw = int(val_hex, 16); return (raw - 0x100000000 if raw >= 0x80000000 else raw) / scale
            elif dtype == 'STRING':
                return binascii.unhexlify(val_hex).decode('gbk', errors='ignore').strip('\x00').strip()
            elif dtype == 'HEX': return val_hex
            return val_hex
        except: return f"ERR:{val_hex}"

    def parse_sub_tlv(self, container_hex, container_id=""):
        res = {}; idx = 0; limit = len(container_hex)
        unknown_subs = []
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
                
                key = f"Sub_{sub_id.upper()}"
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
                else:
                    unknown_subs.append(f"Sub-{container_id}:{sub_id}({val_hex})")
                idx += 4 + (length_val * 2) 
            except: break
        return res, unknown_subs

    def decode_raw(self, raw_hex):
        try:
            raw_str = str(raw_hex).upper().replace("\n", "").replace("\r", "").replace(" ", "")
            hex_blocks = re.findall(r'7E[0-9A-F]{10,}7E', raw_str)
            
            if not hex_blocks:
                return [{'Message ID': 'HEX ERR', '_Status': "❌ No valid JT808 packet found", 'Raw Hex Block': raw_str[:100]}], "Error", 0

            all_decoded_list = []
            final_status = "OK"
            final_mid = 0

            for block in hex_blocks:
                try:
                    raw_bytes = bytes.fromhex(block)
                except ValueError:
                    continue 
                    
                chunks = raw_bytes.split(b'\x7e')
                packets = []
                current_packet = b''
                
                known_msg_ids = {0x0001, 0x0100, 0x0102, 0x0200, 0x0201, 0x0704, 0x0705, 0x0801, 0x0805, 0x0900, 0x8001, 0x8100}
                
                for chunk in chunks:
                    if not chunk:
                        continue
                        
                    is_new_packet = False
                    if len(chunk) >= 12:
                        msg_id = (chunk[0] << 8) | chunk[1]
                        if msg_id in known_msg_ids:
                            is_new_packet = True
                            
                    if is_new_packet:
                        if current_packet:
                            packets.append(current_packet)
                        current_packet = chunk
                    else:
                        if current_packet:
                            current_packet += b'\x7e' + chunk
                        else:
                            current_packet = chunk
                            
                if current_packet:
                    packets.append(current_packet)

                for p_bytes in packets:
                    p_hex = p_bytes.hex().upper()
                    packet_hex = "7E" + p_hex + "7E"
                    
                    data_str = self.unescape(p_hex)
                    try:
                        data_bytes = bytes.fromhex(data_str)
                    except: 
                        all_decoded_list.append({'Message ID': 'HEX ERR', '_Status': "❌ Invalid Hex", 'Raw Hex Block': packet_hex})
                        continue
                        
                    # ป้องกันการ Error ถ้าข้อมูลสั้นเกินไป
                    if len(data_str) < 24:
                        all_decoded_list.append({'Message ID': '?', '_Status': "❌ Packet Too Short", 'Raw Hex Block': packet_hex})
                        continue

                    # [🔥 NEW] ถ้า Checksum พัง แค่แจ้งเตือน แต่ให้อภัยแล้วดึงข้อมูลต่อ!
                    row_status = "OK"
                    if not self.verify_checksum(data_bytes): 
                        row_status = "⚠️ Bad Checksum (Parsed)"

                    data = data_str
                    mid = int(data[0:4], 16)
                    final_mid = mid

                    if mid not in [0x0200, 0x0704]:
                        all_decoded_list.append({
                            'Message ID': f"0x{mid:04X}", 
                            '_Status': "⚠️ System Msg (Ignored)", 
                            'Raw Hex Block': packet_hex
                        })
                        continue

                    body_attr = int(data[4:8], 16); is_2019 = (body_attr >> 14) & 1
                    
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
                        res = {}; std_len_hex = 56 
                        standard_hex = b_hex[:std_len_hex]; ext_hex = b_hex[std_len_hex:]
                        unknowns = []
                        
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
                                
                                if eid in self.container_ids: 
                                    sub_res, unk_subs = self.parse_sub_tlv(e_val, eid)
                                    res.update(sub_res)
                                    unknowns.extend(unk_subs)
                                else:
                                    key = f"Main_{eid.upper()}"
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
                                    else:
                                        unknowns.append(f"Main:{eid}({e_val})")
                                i += 4 + (e_len * 2)
                            except: break

                        for rule in self.calculated_rules:
                            try:
                                target_col = rule['FieldName']
                                ref_col = rule.get('RefField', 'Status')
                                bit_idx = int(rule['BitIndex'])
                                if ref_col in res:
                                    parent_val = int(float(res[ref_col]))
                                    bit_val = (parent_val >> bit_idx) & 1
                                    
                                    t0 = str(rule.get('Text0', '')).strip()
                                    t1 = str(rule.get('Text1', '')).strip()
                                    if t0 == '' and t1 == '':
                                        t0, t1 = "OFF", "ON"
                                        
                                    res[target_col] = t1 if bit_val == 1 else t0
                                else: res[target_col] = "-"
                            except: res[target_col] = "ERR"
                            
                        if unknowns:
                            res['Unknown_Tags'] = " | ".join(unknowns)

                        return {**header, **res, 'Raw Hex Block': packet_hex, '_Status': row_status}

                    body_hex = data[body_start:-2]
                    if mid == 0x0704:
                        try:
                            count = int(body_hex[0:4], 16); curr = 6
                            for _ in range(count):
                                l_len = int(body_hex[curr:curr+4], 16)
                                all_decoded_list.append(parse_one_msg(body_hex[curr+4 : curr+4+(l_len*2)]))
                                curr += 4 + (l_len*2)
                        except: all_decoded_list.append(parse_one_msg(body_hex))
                    else: 
                        all_decoded_list.append(parse_one_msg(body_hex))
            
            if not all_decoded_list:
                return [{'Message ID': 'HEX ERR', '_Status': "❌ No valid JT808 packet found", 'Raw Hex Block': raw_str[:100]}], "Error", 0
                
            return all_decoded_list, final_status, final_mid
            
        except Exception as e: return [{'Message ID': 'CRASH', '_Status': f"❌ Sys Error: {str(e)}"}], "Error", 0