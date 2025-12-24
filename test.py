import pandas as pd

# Creating the master configuration CSV
# Columns: Category, ID, FieldName, StartByte, Length, DataType, Scale
config_data = [
    # Header fields [cite: 37]
    ['Header', None, 'MessageID', 0, 2, 'WORD', 1],
    ['Header', None, 'BodyAttributes', 2, 2, 'WORD', 1],
    ['Header', None, 'TerminalID', 5, 10, 'BCD', 1],
    ['Header', None, 'SerialNumber', 15, 2, 'WORD', 1],

    # Basic Location Information fields 
    ['Location', None, 'AlarmSign', 0, 4, 'DWORD', 1],
    ['Location', None, 'Status', 4, 4, 'DWORD', 1],
    ['Location', None, 'Latitude', 8, 4, 'DWORD', 1000000],
    ['Location', None, 'Longitude', 12, 4, 'DWORD', 1000000],
    ['Location', None, 'Altitude', 16, 2, 'WORD', 1],
    ['Location', None, 'Speed', 18, 2, 'WORD', 10],
    ['Location', None, 'Direction', 20, 2, 'WORD', 1],
    ['Location', None, 'Time', 22, 6, 'BCD', 1],

    # Additional Information Items (TLV) [cite: 130]
    ['Extension', '01', 'Mileage', 0, 4, 'DWORD', 10],
    ['Extension', '30', 'SignalStrength', 0, 1, 'BYTE', 1],
    ['Extension', '31', 'GNSS_Satellites', 0, 1, 'BYTE', 1],

    # BSJ Specific Extensions (Nested in ID EB/EC/ED) 
    ['BSJ_Ext', '002D', 'ExternalVoltage', 0, 4, 'DWORD', 1000],
    ['BSJ_Ext', '00D5', 'IMEI', 0, 15, 'STRING', 1],
    ['BSJ_Ext', '0089', 'ExtendedAlarmStatus', 0, 4, 'DWORD', 1],
    ['BSJ_Ext', '0113', 'FW_Version', 0, 20, 'STRING', 1]
]

df_config = pd.DataFrame(config_data, columns=['Category', 'SubID', 'FieldName', 'StartByte', 'Length', 'DataType', 'Scale'])
df_config.to_csv('master_mapping_config.csv', index=False)