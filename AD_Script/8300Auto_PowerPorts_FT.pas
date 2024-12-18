uses
    SysUtils, Classes, StrUtils;

procedure ImportPowerPortsFromCSV(const CSVFilePath: String);
var
    CSVFile: TextFile;
    Line: String;
    Fields: TStringList;
    SchDoc: ISch_Document;
    PowerPort: ISch_PowerObject;

    PortNamesA1, PortNamesA2, PortNamesA3, PortNamesA4: Array[0..31] of String;
    PortNamesB1, PortNamesB2, PortNamesB3, PortNamesB4: Array[0..31] of String;
    PortNamesC1, PortNamesC2, PortNamesC3, PortNamesC4: Array[0..31] of String;

    PositionA1, PositionA2, PositionA3, PositionA4: Array[0..31] of TPoint;
    PositionB1, PositionB2, PositionB3, PositionB4: Array[0..31] of TPoint;
    PositionC1, PositionC2, PositionC3, PositionC4: Array[0..31] of TPoint;

    i: Integer;
begin
    // 设置 PositionA 的目标位置
    // 初始化 PositionA1 的 32 个坐标
PositionA1[0] := Point(MilsToCoord(1400), MilsToCoord(9400));
PositionA1[1] := Point(MilsToCoord(1400), MilsToCoord(9300));
PositionA1[2] := Point(MilsToCoord(1400), MilsToCoord(9200));
PositionA1[3] := Point(MilsToCoord(1400), MilsToCoord(9100));
PositionA1[4] := Point(MilsToCoord(1400), MilsToCoord(9000));
PositionA1[5] := Point(MilsToCoord(1400), MilsToCoord(8900));
PositionA1[6] := Point(MilsToCoord(1400), MilsToCoord(8800));
PositionA1[7] := Point(MilsToCoord(1400), MilsToCoord(8700));
PositionA1[8] := Point(MilsToCoord(1400), MilsToCoord(8600));
PositionA1[9] := Point(MilsToCoord(1400), MilsToCoord(8500));
PositionA1[10] := Point(MilsToCoord(1400), MilsToCoord(8400));
PositionA1[11] := Point(MilsToCoord(1400), MilsToCoord(8300));
PositionA1[12] := Point(MilsToCoord(1400), MilsToCoord(8200));
PositionA1[13] := Point(MilsToCoord(1400), MilsToCoord(8100));
PositionA1[14] := Point(MilsToCoord(1400), MilsToCoord(8000));
PositionA1[15] := Point(MilsToCoord(1400), MilsToCoord(7900));
PositionA1[16] := Point(MilsToCoord(1400), MilsToCoord(7800));
PositionA1[17] := Point(MilsToCoord(1400), MilsToCoord(7700));
PositionA1[18] := Point(MilsToCoord(1400), MilsToCoord(7600));
PositionA1[19] := Point(MilsToCoord(1400), MilsToCoord(7500));
PositionA1[20] := Point(MilsToCoord(1400), MilsToCoord(7400));
PositionA1[21] := Point(MilsToCoord(1400), MilsToCoord(7300));
PositionA1[22] := Point(MilsToCoord(1400), MilsToCoord(7200));
PositionA1[23] := Point(MilsToCoord(1400), MilsToCoord(7100));
PositionA1[24] := Point(MilsToCoord(1400), MilsToCoord(7000));
PositionA1[25] := Point(MilsToCoord(1400), MilsToCoord(6900));
PositionA1[26] := Point(MilsToCoord(1400), MilsToCoord(6800));
PositionA1[27] := Point(MilsToCoord(1400), MilsToCoord(6700));
PositionA1[28] := Point(MilsToCoord(1400), MilsToCoord(6600));
PositionA1[29] := Point(MilsToCoord(1400), MilsToCoord(6500));
PositionA1[30] := Point(MilsToCoord(1400), MilsToCoord(6400));
PositionA1[31] := Point(MilsToCoord(1400), MilsToCoord(6300));


    // 设置 PositionB 的目标位置
    // 初始化 PositionB1 的 32 个坐标
PositionB1[0] := Point(MilsToCoord(3600), MilsToCoord(9400));
PositionB1[1] := Point(MilsToCoord(3600), MilsToCoord(9300));
PositionB1[2] := Point(MilsToCoord(3600), MilsToCoord(9200));
PositionB1[3] := Point(MilsToCoord(3600), MilsToCoord(9100));
PositionB1[4] := Point(MilsToCoord(3600), MilsToCoord(9000));
PositionB1[5] := Point(MilsToCoord(3600), MilsToCoord(8900));
PositionB1[6] := Point(MilsToCoord(3600), MilsToCoord(8800));
PositionB1[7] := Point(MilsToCoord(3600), MilsToCoord(8700));
PositionB1[8] := Point(MilsToCoord(3600), MilsToCoord(8600));
PositionB1[9] := Point(MilsToCoord(3600), MilsToCoord(8500));
PositionB1[10] := Point(MilsToCoord(3600), MilsToCoord(8400));
PositionB1[11] := Point(MilsToCoord(3600), MilsToCoord(8300));
PositionB1[12] := Point(MilsToCoord(3600), MilsToCoord(8200));
PositionB1[13] := Point(MilsToCoord(3600), MilsToCoord(8100));
PositionB1[14] := Point(MilsToCoord(3600), MilsToCoord(8000));
PositionB1[15] := Point(MilsToCoord(3600), MilsToCoord(7900));
PositionB1[16] := Point(MilsToCoord(3600), MilsToCoord(7800));
PositionB1[17] := Point(MilsToCoord(3600), MilsToCoord(7700));
PositionB1[18] := Point(MilsToCoord(3600), MilsToCoord(7600));
PositionB1[19] := Point(MilsToCoord(3600), MilsToCoord(7500));
PositionB1[20] := Point(MilsToCoord(3600), MilsToCoord(7400));
PositionB1[21] := Point(MilsToCoord(3600), MilsToCoord(7300));
PositionB1[22] := Point(MilsToCoord(3600), MilsToCoord(7200));
PositionB1[23] := Point(MilsToCoord(3600), MilsToCoord(7100));
PositionB1[24] := Point(MilsToCoord(3600), MilsToCoord(7000));
PositionB1[25] := Point(MilsToCoord(3600), MilsToCoord(6900));
PositionB1[26] := Point(MilsToCoord(3600), MilsToCoord(6800));
PositionB1[27] := Point(MilsToCoord(3600), MilsToCoord(6700));
PositionB1[28] := Point(MilsToCoord(3600), MilsToCoord(6600));
PositionB1[29] := Point(MilsToCoord(3600), MilsToCoord(6500));
PositionB1[30] := Point(MilsToCoord(3600), MilsToCoord(6400));
PositionB1[31] := Point(MilsToCoord(3600), MilsToCoord(6300));


// 初始化 PositionC1 的 32 个坐标
PositionC1[0] := Point(MilsToCoord(5800), MilsToCoord(9400));
PositionC1[1] := Point(MilsToCoord(5800), MilsToCoord(9300));
PositionC1[2] := Point(MilsToCoord(5800), MilsToCoord(9200));
PositionC1[3] := Point(MilsToCoord(5800), MilsToCoord(9100));
PositionC1[4] := Point(MilsToCoord(5800), MilsToCoord(9000));
PositionC1[5] := Point(MilsToCoord(5800), MilsToCoord(8900));
PositionC1[6] := Point(MilsToCoord(5800), MilsToCoord(8800));
PositionC1[7] := Point(MilsToCoord(5800), MilsToCoord(8700));
PositionC1[8] := Point(MilsToCoord(5800), MilsToCoord(8600));
PositionC1[9] := Point(MilsToCoord(5800), MilsToCoord(8500));
PositionC1[10] := Point(MilsToCoord(5800), MilsToCoord(8400));
PositionC1[11] := Point(MilsToCoord(5800), MilsToCoord(8300));
PositionC1[12] := Point(MilsToCoord(5800), MilsToCoord(8200));
PositionC1[13] := Point(MilsToCoord(5800), MilsToCoord(8100));
PositionC1[14] := Point(MilsToCoord(5800), MilsToCoord(8000));
PositionC1[15] := Point(MilsToCoord(5800), MilsToCoord(7900));
PositionC1[16] := Point(MilsToCoord(5800), MilsToCoord(7800));
PositionC1[17] := Point(MilsToCoord(5800), MilsToCoord(7700));
PositionC1[18] := Point(MilsToCoord(5800), MilsToCoord(7600));
PositionC1[19] := Point(MilsToCoord(5800), MilsToCoord(7500));
PositionC1[20] := Point(MilsToCoord(5800), MilsToCoord(7400));
PositionC1[21] := Point(MilsToCoord(5800), MilsToCoord(7300));
PositionC1[22] := Point(MilsToCoord(5800), MilsToCoord(7200));
PositionC1[23] := Point(MilsToCoord(5800), MilsToCoord(7100));
PositionC1[24] := Point(MilsToCoord(5800), MilsToCoord(7000));
PositionC1[25] := Point(MilsToCoord(5800), MilsToCoord(6900));
PositionC1[26] := Point(MilsToCoord(5800), MilsToCoord(6800));
PositionC1[27] := Point(MilsToCoord(5800), MilsToCoord(6700));
PositionC1[28] := Point(MilsToCoord(5800), MilsToCoord(6600));
PositionC1[29] := Point(MilsToCoord(5800), MilsToCoord(6500));
PositionC1[30] := Point(MilsToCoord(5800), MilsToCoord(6400));
PositionC1[31] := Point(MilsToCoord(5800), MilsToCoord(6300));


// 初始化 PositionA2 的 32 个坐标
PositionA2[0] := Point(MilsToCoord(8700), MilsToCoord(9400));
PositionA2[1] := Point(MilsToCoord(8700), MilsToCoord(9300));
PositionA2[2] := Point(MilsToCoord(8700), MilsToCoord(9200));
PositionA2[3] := Point(MilsToCoord(8700), MilsToCoord(9100));
PositionA2[4] := Point(MilsToCoord(8700), MilsToCoord(9000));
PositionA2[5] := Point(MilsToCoord(8700), MilsToCoord(8900));
PositionA2[6] := Point(MilsToCoord(8700), MilsToCoord(8800));
PositionA2[7] := Point(MilsToCoord(8700), MilsToCoord(8700));
PositionA2[8] := Point(MilsToCoord(8700), MilsToCoord(8600));
PositionA2[9] := Point(MilsToCoord(8700), MilsToCoord(8500));
PositionA2[10] := Point(MilsToCoord(8700), MilsToCoord(8400));
PositionA2[11] := Point(MilsToCoord(8700), MilsToCoord(8300));
PositionA2[12] := Point(MilsToCoord(8700), MilsToCoord(8200));
PositionA2[13] := Point(MilsToCoord(8700), MilsToCoord(8100));
PositionA2[14] := Point(MilsToCoord(8700), MilsToCoord(8000));
PositionA2[15] := Point(MilsToCoord(8700), MilsToCoord(7900));
PositionA2[16] := Point(MilsToCoord(8700), MilsToCoord(7800));
PositionA2[17] := Point(MilsToCoord(8700), MilsToCoord(7700));
PositionA2[18] := Point(MilsToCoord(8700), MilsToCoord(7600));
PositionA2[19] := Point(MilsToCoord(8700), MilsToCoord(7500));
PositionA2[20] := Point(MilsToCoord(8700), MilsToCoord(7400));
PositionA2[21] := Point(MilsToCoord(8700), MilsToCoord(7300));
PositionA2[22] := Point(MilsToCoord(8700), MilsToCoord(7200));
PositionA2[23] := Point(MilsToCoord(8700), MilsToCoord(7100));
PositionA2[24] := Point(MilsToCoord(8700), MilsToCoord(7000));
PositionA2[25] := Point(MilsToCoord(8700), MilsToCoord(6900));
PositionA2[26] := Point(MilsToCoord(8700), MilsToCoord(6800));
PositionA2[27] := Point(MilsToCoord(8700), MilsToCoord(6700));
PositionA2[28] := Point(MilsToCoord(8700), MilsToCoord(6600));
PositionA2[29] := Point(MilsToCoord(8700), MilsToCoord(6500));
PositionA2[30] := Point(MilsToCoord(8700), MilsToCoord(6400));
PositionA2[31] := Point(MilsToCoord(8700), MilsToCoord(6300));


    // 初始化 PositionB2 的 32 个坐标
PositionB2[0] := Point(MilsToCoord(10900), MilsToCoord(9400));
PositionB2[1] := Point(MilsToCoord(10900), MilsToCoord(9300));
PositionB2[2] := Point(MilsToCoord(10900), MilsToCoord(9200));
PositionB2[3] := Point(MilsToCoord(10900), MilsToCoord(9100));
PositionB2[4] := Point(MilsToCoord(10900), MilsToCoord(9000));
PositionB2[5] := Point(MilsToCoord(10900), MilsToCoord(8900));
PositionB2[6] := Point(MilsToCoord(10900), MilsToCoord(8800));
PositionB2[7] := Point(MilsToCoord(10900), MilsToCoord(8700));
PositionB2[8] := Point(MilsToCoord(10900), MilsToCoord(8600));
PositionB2[9] := Point(MilsToCoord(10900), MilsToCoord(8500));
PositionB2[10] := Point(MilsToCoord(10900), MilsToCoord(8400));
PositionB2[11] := Point(MilsToCoord(10900), MilsToCoord(8300));
PositionB2[12] := Point(MilsToCoord(10900), MilsToCoord(8200));
PositionB2[13] := Point(MilsToCoord(10900), MilsToCoord(8100));
PositionB2[14] := Point(MilsToCoord(10900), MilsToCoord(8000));
PositionB2[15] := Point(MilsToCoord(10900), MilsToCoord(7900));
PositionB2[16] := Point(MilsToCoord(10900), MilsToCoord(7800));
PositionB2[17] := Point(MilsToCoord(10900), MilsToCoord(7700));
PositionB2[18] := Point(MilsToCoord(10900), MilsToCoord(7600));
PositionB2[19] := Point(MilsToCoord(10900), MilsToCoord(7500));
PositionB2[20] := Point(MilsToCoord(10900), MilsToCoord(7400));
PositionB2[21] := Point(MilsToCoord(10900), MilsToCoord(7300));
PositionB2[22] := Point(MilsToCoord(10900), MilsToCoord(7200));
PositionB2[23] := Point(MilsToCoord(10900), MilsToCoord(7100));
PositionB2[24] := Point(MilsToCoord(10900), MilsToCoord(7000));
PositionB2[25] := Point(MilsToCoord(10900), MilsToCoord(6900));
PositionB2[26] := Point(MilsToCoord(10900), MilsToCoord(6800));
PositionB2[27] := Point(MilsToCoord(10900), MilsToCoord(6700));
PositionB2[28] := Point(MilsToCoord(10900), MilsToCoord(6600));
PositionB2[29] := Point(MilsToCoord(10900), MilsToCoord(6500));
PositionB2[30] := Point(MilsToCoord(10900), MilsToCoord(6400));
PositionB2[31] := Point(MilsToCoord(10900), MilsToCoord(6300));

    // 初始化 PositionC2 的 32 个坐标
PositionC2[0] := Point(MilsToCoord(13100), MilsToCoord(9400));
PositionC2[1] := Point(MilsToCoord(13100), MilsToCoord(9300));
PositionC2[2] := Point(MilsToCoord(13100), MilsToCoord(9200));
PositionC2[3] := Point(MilsToCoord(13100), MilsToCoord(9100));
PositionC2[4] := Point(MilsToCoord(13100), MilsToCoord(9000));
PositionC2[5] := Point(MilsToCoord(13100), MilsToCoord(8900));
PositionC2[6] := Point(MilsToCoord(13100), MilsToCoord(8800));
PositionC2[7] := Point(MilsToCoord(13100), MilsToCoord(8700));
PositionC2[8] := Point(MilsToCoord(13100), MilsToCoord(8600));
PositionC2[9] := Point(MilsToCoord(13100), MilsToCoord(8500));
PositionC2[10] := Point(MilsToCoord(13100), MilsToCoord(8400));
PositionC2[11] := Point(MilsToCoord(13100), MilsToCoord(8300));
PositionC2[12] := Point(MilsToCoord(13100), MilsToCoord(8200));
PositionC2[13] := Point(MilsToCoord(13100), MilsToCoord(8100));
PositionC2[14] := Point(MilsToCoord(13100), MilsToCoord(8000));
PositionC2[15] := Point(MilsToCoord(13100), MilsToCoord(7900));
PositionC2[16] := Point(MilsToCoord(13100), MilsToCoord(7800));
PositionC2[17] := Point(MilsToCoord(13100), MilsToCoord(7700));
PositionC2[18] := Point(MilsToCoord(13100), MilsToCoord(7600));
PositionC2[19] := Point(MilsToCoord(13100), MilsToCoord(7500));
PositionC2[20] := Point(MilsToCoord(13100), MilsToCoord(7400));
PositionC2[21] := Point(MilsToCoord(13100), MilsToCoord(7300));
PositionC2[22] := Point(MilsToCoord(13100), MilsToCoord(7200));
PositionC2[23] := Point(MilsToCoord(13100), MilsToCoord(7100));
PositionC2[24] := Point(MilsToCoord(13100), MilsToCoord(7000));
PositionC2[25] := Point(MilsToCoord(13100), MilsToCoord(6900));
PositionC2[26] := Point(MilsToCoord(13100), MilsToCoord(6800));
PositionC2[27] := Point(MilsToCoord(13100), MilsToCoord(6700));
PositionC2[28] := Point(MilsToCoord(13100), MilsToCoord(6600));
PositionC2[29] := Point(MilsToCoord(13100), MilsToCoord(6500));
PositionC2[30] := Point(MilsToCoord(13100), MilsToCoord(6400));
PositionC2[31] := Point(MilsToCoord(13100), MilsToCoord(6300));


    // 初始化 PositionA3 的 32 个坐标
PositionA3[0] := Point(MilsToCoord(1400), MilsToCoord(4300));
PositionA3[1] := Point(MilsToCoord(1400), MilsToCoord(4200));
PositionA3[2] := Point(MilsToCoord(1400), MilsToCoord(4100));
PositionA3[3] := Point(MilsToCoord(1400), MilsToCoord(4000));
PositionA3[4] := Point(MilsToCoord(1400), MilsToCoord(3900));
PositionA3[5] := Point(MilsToCoord(1400), MilsToCoord(3800));
PositionA3[6] := Point(MilsToCoord(1400), MilsToCoord(3700));
PositionA3[7] := Point(MilsToCoord(1400), MilsToCoord(3600));
PositionA3[8] := Point(MilsToCoord(1400), MilsToCoord(3500));
PositionA3[9] := Point(MilsToCoord(1400), MilsToCoord(3400));
PositionA3[10] := Point(MilsToCoord(1400), MilsToCoord(3300));
PositionA3[11] := Point(MilsToCoord(1400), MilsToCoord(3200));
PositionA3[12] := Point(MilsToCoord(1400), MilsToCoord(3100));
PositionA3[13] := Point(MilsToCoord(1400), MilsToCoord(3000));
PositionA3[14] := Point(MilsToCoord(1400), MilsToCoord(2900));
PositionA3[15] := Point(MilsToCoord(1400), MilsToCoord(2800));
PositionA3[16] := Point(MilsToCoord(1400), MilsToCoord(2700));
PositionA3[17] := Point(MilsToCoord(1400), MilsToCoord(2600));
PositionA3[18] := Point(MilsToCoord(1400), MilsToCoord(2500));
PositionA3[19] := Point(MilsToCoord(1400), MilsToCoord(2400));
PositionA3[20] := Point(MilsToCoord(1400), MilsToCoord(2300));
PositionA3[21] := Point(MilsToCoord(1400), MilsToCoord(2200));
PositionA3[22] := Point(MilsToCoord(1400), MilsToCoord(2100));
PositionA3[23] := Point(MilsToCoord(1400), MilsToCoord(2000));
PositionA3[24] := Point(MilsToCoord(1400), MilsToCoord(1900));
PositionA3[25] := Point(MilsToCoord(1400), MilsToCoord(1800));
PositionA3[26] := Point(MilsToCoord(1400), MilsToCoord(1700));
PositionA3[27] := Point(MilsToCoord(1400), MilsToCoord(1600));
PositionA3[28] := Point(MilsToCoord(1400), MilsToCoord(1500));
PositionA3[29] := Point(MilsToCoord(1400), MilsToCoord(1400));
PositionA3[30] := Point(MilsToCoord(1400), MilsToCoord(1300));
PositionA3[31] := Point(MilsToCoord(1400), MilsToCoord(1200));


// 初始化 PositionB3 的 32 个坐标
PositionB3[0] := Point(MilsToCoord(3600), MilsToCoord(4300));
PositionB3[1] := Point(MilsToCoord(3600), MilsToCoord(4200));
PositionB3[2] := Point(MilsToCoord(3600), MilsToCoord(4100));
PositionB3[3] := Point(MilsToCoord(3600), MilsToCoord(4000));
PositionB3[4] := Point(MilsToCoord(3600), MilsToCoord(3900));
PositionB3[5] := Point(MilsToCoord(3600), MilsToCoord(3800));
PositionB3[6] := Point(MilsToCoord(3600), MilsToCoord(3700));
PositionB3[7] := Point(MilsToCoord(3600), MilsToCoord(3600));
PositionB3[8] := Point(MilsToCoord(3600), MilsToCoord(3500));
PositionB3[9] := Point(MilsToCoord(3600), MilsToCoord(3400));
PositionB3[10] := Point(MilsToCoord(3600), MilsToCoord(3300));
PositionB3[11] := Point(MilsToCoord(3600), MilsToCoord(3200));
PositionB3[12] := Point(MilsToCoord(3600), MilsToCoord(3100));
PositionB3[13] := Point(MilsToCoord(3600), MilsToCoord(3000));
PositionB3[14] := Point(MilsToCoord(3600), MilsToCoord(2900));
PositionB3[15] := Point(MilsToCoord(3600), MilsToCoord(2800));
PositionB3[16] := Point(MilsToCoord(3600), MilsToCoord(2700));
PositionB3[17] := Point(MilsToCoord(3600), MilsToCoord(2600));
PositionB3[18] := Point(MilsToCoord(3600), MilsToCoord(2500));
PositionB3[19] := Point(MilsToCoord(3600), MilsToCoord(2400));
PositionB3[20] := Point(MilsToCoord(3600), MilsToCoord(2300));
PositionB3[21] := Point(MilsToCoord(3600), MilsToCoord(2200));
PositionB3[22] := Point(MilsToCoord(3600), MilsToCoord(2100));
PositionB3[23] := Point(MilsToCoord(3600), MilsToCoord(2000));
PositionB3[24] := Point(MilsToCoord(3600), MilsToCoord(1900));
PositionB3[25] := Point(MilsToCoord(3600), MilsToCoord(1800));
PositionB3[26] := Point(MilsToCoord(3600), MilsToCoord(1700));
PositionB3[27] := Point(MilsToCoord(3600), MilsToCoord(1600));
PositionB3[28] := Point(MilsToCoord(3600), MilsToCoord(1500));
PositionB3[29] := Point(MilsToCoord(3600), MilsToCoord(1400));
PositionB3[30] := Point(MilsToCoord(3600), MilsToCoord(1300));
PositionB3[31] := Point(MilsToCoord(3600), MilsToCoord(1200));


// 初始化 PositionC3 的 32 个坐标
PositionC3[0] := Point(MilsToCoord(5800), MilsToCoord(4300));
PositionC3[1] := Point(MilsToCoord(5800), MilsToCoord(4200));
PositionC3[2] := Point(MilsToCoord(5800), MilsToCoord(4100));
PositionC3[3] := Point(MilsToCoord(5800), MilsToCoord(4000));
PositionC3[4] := Point(MilsToCoord(5800), MilsToCoord(3900));
PositionC3[5] := Point(MilsToCoord(5800), MilsToCoord(3800));
PositionC3[6] := Point(MilsToCoord(5800), MilsToCoord(3700));
PositionC3[7] := Point(MilsToCoord(5800), MilsToCoord(3600));
PositionC3[8] := Point(MilsToCoord(5800), MilsToCoord(3500));
PositionC3[9] := Point(MilsToCoord(5800), MilsToCoord(3400));
PositionC3[10] := Point(MilsToCoord(5800), MilsToCoord(3300));
PositionC3[11] := Point(MilsToCoord(5800), MilsToCoord(3200));
PositionC3[12] := Point(MilsToCoord(5800), MilsToCoord(3100));
PositionC3[13] := Point(MilsToCoord(5800), MilsToCoord(3000));
PositionC3[14] := Point(MilsToCoord(5800), MilsToCoord(2900));
PositionC3[15] := Point(MilsToCoord(5800), MilsToCoord(2800));
PositionC3[16] := Point(MilsToCoord(5800), MilsToCoord(2700));
PositionC3[17] := Point(MilsToCoord(5800), MilsToCoord(2600));
PositionC3[18] := Point(MilsToCoord(5800), MilsToCoord(2500));
PositionC3[19] := Point(MilsToCoord(5800), MilsToCoord(2400));
PositionC3[20] := Point(MilsToCoord(5800), MilsToCoord(2300));
PositionC3[21] := Point(MilsToCoord(5800), MilsToCoord(2200));
PositionC3[22] := Point(MilsToCoord(5800), MilsToCoord(2100));
PositionC3[23] := Point(MilsToCoord(5800), MilsToCoord(2000));
PositionC3[24] := Point(MilsToCoord(5800), MilsToCoord(1900));
PositionC3[25] := Point(MilsToCoord(5800), MilsToCoord(1800));
PositionC3[26] := Point(MilsToCoord(5800), MilsToCoord(1700));
PositionC3[27] := Point(MilsToCoord(5800), MilsToCoord(1600));
PositionC3[28] := Point(MilsToCoord(5800), MilsToCoord(1500));
PositionC3[29] := Point(MilsToCoord(5800), MilsToCoord(1400));
PositionC3[30] := Point(MilsToCoord(5800), MilsToCoord(1300));
PositionC3[31] := Point(MilsToCoord(5800), MilsToCoord(1200));


// 初始化 PositionA4 的 32 个坐标
PositionA4[0] := Point(MilsToCoord(8700), MilsToCoord(4300));
PositionA4[1] := Point(MilsToCoord(8700), MilsToCoord(4200));
PositionA4[2] := Point(MilsToCoord(8700), MilsToCoord(4100));
PositionA4[3] := Point(MilsToCoord(8700), MilsToCoord(4000));
PositionA4[4] := Point(MilsToCoord(8700), MilsToCoord(3900));
PositionA4[5] := Point(MilsToCoord(8700), MilsToCoord(3800));
PositionA4[6] := Point(MilsToCoord(8700), MilsToCoord(3700));
PositionA4[7] := Point(MilsToCoord(8700), MilsToCoord(3600));
PositionA4[8] := Point(MilsToCoord(8700), MilsToCoord(3500));
PositionA4[9] := Point(MilsToCoord(8700), MilsToCoord(3400));
PositionA4[10] := Point(MilsToCoord(8700), MilsToCoord(3300));
PositionA4[11] := Point(MilsToCoord(8700), MilsToCoord(3200));
PositionA4[12] := Point(MilsToCoord(8700), MilsToCoord(3100));
PositionA4[13] := Point(MilsToCoord(8700), MilsToCoord(3000));
PositionA4[14] := Point(MilsToCoord(8700), MilsToCoord(2900));
PositionA4[15] := Point(MilsToCoord(8700), MilsToCoord(2800));
PositionA4[16] := Point(MilsToCoord(8700), MilsToCoord(2700));
PositionA4[17] := Point(MilsToCoord(8700), MilsToCoord(2600));
PositionA4[18] := Point(MilsToCoord(8700), MilsToCoord(2500));
PositionA4[19] := Point(MilsToCoord(8700), MilsToCoord(2400));
PositionA4[20] := Point(MilsToCoord(8700), MilsToCoord(2300));
PositionA4[21] := Point(MilsToCoord(8700), MilsToCoord(2200));
PositionA4[22] := Point(MilsToCoord(8700), MilsToCoord(2100));
PositionA4[23] := Point(MilsToCoord(8700), MilsToCoord(2000));
PositionA4[24] := Point(MilsToCoord(8700), MilsToCoord(1900));
PositionA4[25] := Point(MilsToCoord(8700), MilsToCoord(1800));
PositionA4[26] := Point(MilsToCoord(8700), MilsToCoord(1700));
PositionA4[27] := Point(MilsToCoord(8700), MilsToCoord(1600));
PositionA4[28] := Point(MilsToCoord(8700), MilsToCoord(1500));
PositionA4[29] := Point(MilsToCoord(8700), MilsToCoord(1400));
PositionA4[30] := Point(MilsToCoord(8700), MilsToCoord(1300));
PositionA4[31] := Point(MilsToCoord(8700), MilsToCoord(1200));


// 初始化 PositionB4 的 32 个坐标
PositionB4[0] := Point(MilsToCoord(10900), MilsToCoord(4300));
PositionB4[1] := Point(MilsToCoord(10900), MilsToCoord(4200));
PositionB4[2] := Point(MilsToCoord(10900), MilsToCoord(4100));
PositionB4[3] := Point(MilsToCoord(10900), MilsToCoord(4000));
PositionB4[4] := Point(MilsToCoord(10900), MilsToCoord(3900));
PositionB4[5] := Point(MilsToCoord(10900), MilsToCoord(3800));
PositionB4[6] := Point(MilsToCoord(10900), MilsToCoord(3700));
PositionB4[7] := Point(MilsToCoord(10900), MilsToCoord(3600));
PositionB4[8] := Point(MilsToCoord(10900), MilsToCoord(3500));
PositionB4[9] := Point(MilsToCoord(10900), MilsToCoord(3400));
PositionB4[10] := Point(MilsToCoord(10900), MilsToCoord(3300));
PositionB4[11] := Point(MilsToCoord(10900), MilsToCoord(3200));
PositionB4[12] := Point(MilsToCoord(10900), MilsToCoord(3100));
PositionB4[13] := Point(MilsToCoord(10900), MilsToCoord(3000));
PositionB4[14] := Point(MilsToCoord(10900), MilsToCoord(2900));
PositionB4[15] := Point(MilsToCoord(10900), MilsToCoord(2800));
PositionB4[16] := Point(MilsToCoord(10900), MilsToCoord(2700));
PositionB4[17] := Point(MilsToCoord(10900), MilsToCoord(2600));
PositionB4[18] := Point(MilsToCoord(10900), MilsToCoord(2500));
PositionB4[19] := Point(MilsToCoord(10900), MilsToCoord(2400));
PositionB4[20] := Point(MilsToCoord(10900), MilsToCoord(2300));
PositionB4[21] := Point(MilsToCoord(10900), MilsToCoord(2200));
PositionB4[22] := Point(MilsToCoord(10900), MilsToCoord(2100));
PositionB4[23] := Point(MilsToCoord(10900), MilsToCoord(2000));
PositionB4[24] := Point(MilsToCoord(10900), MilsToCoord(1900));
PositionB4[25] := Point(MilsToCoord(10900), MilsToCoord(1800));
PositionB4[26] := Point(MilsToCoord(10900), MilsToCoord(1700));
PositionB4[27] := Point(MilsToCoord(10900), MilsToCoord(1600));
PositionB4[28] := Point(MilsToCoord(10900), MilsToCoord(1500));
PositionB4[29] := Point(MilsToCoord(10900), MilsToCoord(1400));
PositionB4[30] := Point(MilsToCoord(10900), MilsToCoord(1300));
PositionB4[31] := Point(MilsToCoord(10900), MilsToCoord(1200));


// 初始化 PositionC4 的 32 个坐标
PositionC4[0] := Point(MilsToCoord(13100), MilsToCoord(4300));
PositionC4[1] := Point(MilsToCoord(13100), MilsToCoord(4200));
PositionC4[2] := Point(MilsToCoord(13100), MilsToCoord(4100));
PositionC4[3] := Point(MilsToCoord(13100), MilsToCoord(4000));
PositionC4[4] := Point(MilsToCoord(13100), MilsToCoord(3900));
PositionC4[5] := Point(MilsToCoord(13100), MilsToCoord(3800));
PositionC4[6] := Point(MilsToCoord(13100), MilsToCoord(3700));
PositionC4[7] := Point(MilsToCoord(13100), MilsToCoord(3600));
PositionC4[8] := Point(MilsToCoord(13100), MilsToCoord(3500));
PositionC4[9] := Point(MilsToCoord(13100), MilsToCoord(3400));
PositionC4[10] := Point(MilsToCoord(13100), MilsToCoord(3300));
PositionC4[11] := Point(MilsToCoord(13100), MilsToCoord(3200));
PositionC4[12] := Point(MilsToCoord(13100), MilsToCoord(3100));
PositionC4[13] := Point(MilsToCoord(13100), MilsToCoord(3000));
PositionC4[14] := Point(MilsToCoord(13100), MilsToCoord(2900));
PositionC4[15] := Point(MilsToCoord(13100), MilsToCoord(2800));
PositionC4[16] := Point(MilsToCoord(13100), MilsToCoord(2700));
PositionC4[17] := Point(MilsToCoord(13100), MilsToCoord(2600));
PositionC4[18] := Point(MilsToCoord(13100), MilsToCoord(2500));
PositionC4[19] := Point(MilsToCoord(13100), MilsToCoord(2400));
PositionC4[20] := Point(MilsToCoord(13100), MilsToCoord(2300));
PositionC4[21] := Point(MilsToCoord(13100), MilsToCoord(2200));
PositionC4[22] := Point(MilsToCoord(13100), MilsToCoord(2100));
PositionC4[23] := Point(MilsToCoord(13100), MilsToCoord(2000));
PositionC4[24] := Point(MilsToCoord(13100), MilsToCoord(1900));
PositionC4[25] := Point(MilsToCoord(13100), MilsToCoord(1800));
PositionC4[26] := Point(MilsToCoord(13100), MilsToCoord(1700));
PositionC4[27] := Point(MilsToCoord(13100), MilsToCoord(1600));
PositionC4[28] := Point(MilsToCoord(13100), MilsToCoord(1500));
PositionC4[29] := Point(MilsToCoord(13100), MilsToCoord(1400));
PositionC4[30] := Point(MilsToCoord(13100), MilsToCoord(1300));
PositionC4[31] := Point(MilsToCoord(13100), MilsToCoord(1200));






    // 获取当前打开的原理图文档
    SchDoc := SchServer.GetCurrentSchDocument;
    if SchDoc = nil then
    begin
        ShowMessage('请打开一个原理图文档！');
        Exit;
    end;

    // 打开 CSV 文件
    AssignFile(CSVFile, CSVFilePath);
    Reset(CSVFile);

    try
        // 跳过第一行标题
        Readln(CSVFile, Line);

        // 读取数据行并存入 PortNamesA 和 PortNamesB 数组
        i := 0;
        while (not EOF(CSVFile)) and (i < 32) do
        begin
            Readln(CSVFile, Line);
            Fields := TStringList.Create;
            Fields.Delimiter := ',';
            Fields.StrictDelimiter := True;
            Fields.DelimitedText := Line;

            // 取第一列和第二列的数据
            if Fields.Count > 1 then
            begin
               PortNamesA1[i] := Fields[0]; // 将第一列的数据存入PortNamesA1数组
               PortNamesB1[i] := Fields[1]; // 将第二列的数据存入PortNamesB1数组
               PortNamesC1[i] := Fields[2];
               PortNamesA2[i] := Fields[3];
               PortNamesB2[i] := Fields[4];
               PortNamesC2[i] := Fields[5];
               PortNamesA3[i] := Fields[6];
               PortNamesB3[i] := Fields[7];
               PortNamesC3[i] := Fields[8];
               PortNamesA4[i] := Fields[9];
               PortNamesB4[i] := Fields[10];
               PortNamesC4[i] := Fields[11];



            end;

            Fields.Free;
            Inc(i);
        end;


// 在 A1 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesA1[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionA1[i];
            PowerPort.Text := PortNamesA1[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 B1 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesB1[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionB1[i];
            PowerPort.Text := PortNamesB1[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 C1 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesC1[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionC1[i];
            PowerPort.Text := PortNamesC1[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 A2 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesA2[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionA2[i];
            PowerPort.Text := PortNamesA2[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 B2 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesB2[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionB2[i];
            PowerPort.Text := PortNamesB2[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 C2 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesC2[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionC2[i];
            PowerPort.Text := PortNamesC2[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 A3 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesA3[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionA3[i];
            PowerPort.Text := PortNamesA3[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 B3 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesB3[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionB3[i];
            PowerPort.Text := PortNamesB3[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 C3 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesC3[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionC3[i];
            PowerPort.Text := PortNamesC3[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 A4 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesA4[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionA4[i];
            PowerPort.Text := PortNamesA4[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 B4 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesB4[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionB4[i];
            PowerPort.Text := PortNamesB4[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;

// 在 C4 位置生成 Power Ports
for i := 0 to 31 do
begin
    if PortNamesC4[i] <> '' then  // 检查是否为空
    begin
        PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
        if PowerPort <> nil then
        begin
            PowerPort.Location := PositionC4[i];
            PowerPort.Text := PortNamesC4[i];

            SchDoc.RegisterSchObjectInContainer(PowerPort);
            SchDoc.AddSchObject(PowerPort);
        end;
    end;
end;



    finally
        // 关闭 CSV 文件
        CloseFile(CSVFile);
    end;

    ShowMessage('Power ports 已成功从 CSV 导入到原理图。');
end;


// 定义 8 个独立的函数
procedure ImportPowerPorts_Connector_X1X4;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\8300Auto_PowerPorts_demo_V1.2\Connector_X1-4.csv');
end;

procedure ImportPowerPorts_Connector_X5X8;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\8300Auto_PowerPorts_demo_V1.2\Connector_X5-8.csv');
end;

procedure ImportPowerPorts_Connector_X9X12;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\8300Auto_PowerPorts_demo_V1.2\Connector_X9-12.csv');
end;

procedure ImportPowerPorts_Connector_X13X16;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\8300Auto_PowerPorts_demo_V1.2\Connector_X13-16.csv');
end;























