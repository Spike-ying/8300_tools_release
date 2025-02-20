// AD Script for CP Test Octant 1~8 PowerPorts
//Please refill the right CSV files path in the bottom of the codes
//Bug report: Spike.ying@accotest.com
uses
    SysUtils, Classes, StrUtils;

procedure ImportPowerPortsFromCSV(const CSVFilePath: String);
var
    CSVFile: TextFile;
    Line: String;
    Fields: TStringList;
    SchDoc: ISch_Document;
    PowerPort: ISch_PowerObject;
    PortNamesA: Array[0..21] of String; // 用于存储第一列的22个端口名称
    PortNamesB: Array[0..21] of String; // 用于存储第二列的22个端口名称
    PortNamesC: Array[0..21] of String;
    PortNamesD: Array[0..21] of String;
    PortNamesE: Array[0..21] of String;
    PortNamesF: Array[0..21] of String;
    PortNamesG: Array[0..21] of String;
    PortNamesH: Array[0..21] of String;
    PortNamesI: Array[0..21] of String;
    PortNamesJ: Array[0..21] of String;
    PortNamesK: Array[0..21] of String;
    PortNamesL: Array[0..21] of String;
    PortNamesM: Array[0..21] of String;
    PortNamesN: Array[0..21] of String;
    PortNamesO: Array[0..21] of String;
    PortNamesP: Array[0..21] of String;
    PortNamesQ: Array[0..21] of String;

    PositionA: Array[0..21] of TPoint;  // 用于存储PositionA的22个位置
    PositionB: Array[0..21] of TPoint;  // 用于存储PositionB的22个位置
    PositionC: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionD: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionE: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionF: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionG: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionH: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionI: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionJ: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionK: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionL: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionM: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionN: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionO: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionP: Array[0..21] of TPoint; // PositionC 的 22 个位置
    PositionQ: Array[0..21] of TPoint; // PositionC 的 22 个位置
    i: Integer;
begin
    // 设置 PositionA 的目标位置
    PositionA[0] := Point(MilsToCoord(800), MilsToCoord(10600));
    PositionA[1] := Point(MilsToCoord(800), MilsToCoord(10500));
    PositionA[2] := Point(MilsToCoord(800), MilsToCoord(10400));
    PositionA[3] := Point(MilsToCoord(800), MilsToCoord(10300));
    PositionA[4] := Point(MilsToCoord(800), MilsToCoord(10200));
    PositionA[5] := Point(MilsToCoord(800), MilsToCoord(10100));
    PositionA[6] := Point(MilsToCoord(800), MilsToCoord(10000));
    PositionA[7] := Point(MilsToCoord(800), MilsToCoord(9900));
    PositionA[8] := Point(MilsToCoord(800), MilsToCoord(9800));
    PositionA[9] := Point(MilsToCoord(800), MilsToCoord(9700));
    PositionA[10] := Point(MilsToCoord(800), MilsToCoord(9600));
    PositionA[11] := Point(MilsToCoord(800), MilsToCoord(9500));
    PositionA[12] := Point(MilsToCoord(800), MilsToCoord(9400));
    PositionA[13] := Point(MilsToCoord(800), MilsToCoord(9300));
    PositionA[14] := Point(MilsToCoord(800), MilsToCoord(9200));
    PositionA[15] := Point(MilsToCoord(800), MilsToCoord(9100));
    PositionA[16] := Point(MilsToCoord(800), MilsToCoord(9000));
    PositionA[17] := Point(MilsToCoord(800), MilsToCoord(8900));
    PositionA[18] := Point(MilsToCoord(800), MilsToCoord(8800));
    PositionA[19] := Point(MilsToCoord(800), MilsToCoord(8700));
    PositionA[20] := Point(MilsToCoord(800), MilsToCoord(8600));
    PositionA[21] := Point(MilsToCoord(800), MilsToCoord(8500));

    // 设置 PositionB 的目标位置
    PositionB[0] := Point(MilsToCoord(2500), MilsToCoord(10600));
    PositionB[1] := Point(MilsToCoord(2500), MilsToCoord(10500));
    PositionB[2] := Point(MilsToCoord(2500), MilsToCoord(10400));
    PositionB[3] := Point(MilsToCoord(2500), MilsToCoord(10300));
    PositionB[4] := Point(MilsToCoord(2500), MilsToCoord(10200));
    PositionB[5] := Point(MilsToCoord(2500), MilsToCoord(10100));
    PositionB[6] := Point(MilsToCoord(2500), MilsToCoord(10000));
    PositionB[7] := Point(MilsToCoord(2500), MilsToCoord(9900));
    PositionB[8] := Point(MilsToCoord(2500), MilsToCoord(9800));
    PositionB[9] := Point(MilsToCoord(2500), MilsToCoord(9700));
    PositionB[10] := Point(MilsToCoord(2500), MilsToCoord(9600));
    PositionB[11] := Point(MilsToCoord(2500), MilsToCoord(9500));
    PositionB[12] := Point(MilsToCoord(2500), MilsToCoord(9400));
    PositionB[13] := Point(MilsToCoord(2500), MilsToCoord(9300));
    PositionB[14] := Point(MilsToCoord(2500), MilsToCoord(9200));
    PositionB[15] := Point(MilsToCoord(2500), MilsToCoord(9100));
    PositionB[16] := Point(MilsToCoord(2500), MilsToCoord(9000));
    PositionB[17] := Point(MilsToCoord(2500), MilsToCoord(8900));
    PositionB[18] := Point(MilsToCoord(2500), MilsToCoord(8800));
    PositionB[19] := Point(MilsToCoord(2500), MilsToCoord(8700));
    PositionB[20] := Point(MilsToCoord(2500), MilsToCoord(8600));
    PositionB[21] := Point(MilsToCoord(2500), MilsToCoord(8500));

    // 设置 PositionC 的目标位置
    PositionC[0] := Point(MilsToCoord(3900), MilsToCoord(10600));
    PositionC[1] := Point(MilsToCoord(3900), MilsToCoord(10500));
    PositionC[2] := Point(MilsToCoord(3900), MilsToCoord(10400));
    PositionC[3] := Point(MilsToCoord(3900), MilsToCoord(10300));
    PositionC[4] := Point(MilsToCoord(3900), MilsToCoord(10200));
    PositionC[5] := Point(MilsToCoord(3900), MilsToCoord(10100));
    PositionC[6] := Point(MilsToCoord(3900), MilsToCoord(10000));
    PositionC[7] := Point(MilsToCoord(3900), MilsToCoord(9900));
    PositionC[8] := Point(MilsToCoord(3900), MilsToCoord(9800));
    PositionC[9] := Point(MilsToCoord(3900), MilsToCoord(9700));
    PositionC[10] := Point(MilsToCoord(3900), MilsToCoord(9600));
    PositionC[11] := Point(MilsToCoord(3900), MilsToCoord(9500));
    PositionC[12] := Point(MilsToCoord(3900), MilsToCoord(9400));
    PositionC[13] := Point(MilsToCoord(3900), MilsToCoord(9300));
    PositionC[14] := Point(MilsToCoord(3900), MilsToCoord(9200));
    PositionC[15] := Point(MilsToCoord(3900), MilsToCoord(9100));
    PositionC[16] := Point(MilsToCoord(3900), MilsToCoord(9000));
    PositionC[17] := Point(MilsToCoord(3900), MilsToCoord(8900));
    PositionC[18] := Point(MilsToCoord(3900), MilsToCoord(8800));
    PositionC[19] := Point(MilsToCoord(3900), MilsToCoord(8700));
    PositionC[20] := Point(MilsToCoord(3900), MilsToCoord(8600));
    PositionC[21] := Point(MilsToCoord(3900), MilsToCoord(8500));

    //D
    PositionD[0] := Point(MilsToCoord(5600), MilsToCoord(10600));
    PositionD[1] := Point(MilsToCoord(5600), MilsToCoord(10500));
    PositionD[2] := Point(MilsToCoord(5600), MilsToCoord(10400));
    PositionD[3] := Point(MilsToCoord(5600), MilsToCoord(10300));
    PositionD[4] := Point(MilsToCoord(5600), MilsToCoord(10200));
    PositionD[5] := Point(MilsToCoord(5600), MilsToCoord(10100));
    PositionD[6] := Point(MilsToCoord(5600), MilsToCoord(10000));
    PositionD[7] := Point(MilsToCoord(5600), MilsToCoord(9900));
    PositionD[8] := Point(MilsToCoord(5600), MilsToCoord(9800));
    PositionD[9] := Point(MilsToCoord(5600), MilsToCoord(9700));
    PositionD[10] := Point(MilsToCoord(5600), MilsToCoord(9600));
    PositionD[11] := Point(MilsToCoord(5600), MilsToCoord(9500));
    PositionD[12] := Point(MilsToCoord(5600), MilsToCoord(9400));
    PositionD[13] := Point(MilsToCoord(5600), MilsToCoord(9300));
    PositionD[14] := Point(MilsToCoord(5600), MilsToCoord(9200));
    PositionD[15] := Point(MilsToCoord(5600), MilsToCoord(9100));
    PositionD[16] := Point(MilsToCoord(5600), MilsToCoord(9000));
    PositionD[17] := Point(MilsToCoord(5600), MilsToCoord(8900));
    PositionD[18] := Point(MilsToCoord(5600), MilsToCoord(8800));
    PositionD[19] := Point(MilsToCoord(5600), MilsToCoord(8700));
    PositionD[20] := Point(MilsToCoord(5600), MilsToCoord(8600));
    PositionD[21] := Point(MilsToCoord(5600), MilsToCoord(8500));

    //E
    PositionE[0] := Point(MilsToCoord(7100), MilsToCoord(10600));
    PositionE[1] := Point(MilsToCoord(7100), MilsToCoord(10500));
    PositionE[2] := Point(MilsToCoord(7100), MilsToCoord(10400));
    PositionE[3] := Point(MilsToCoord(7100), MilsToCoord(10300));
    PositionE[4] := Point(MilsToCoord(7100), MilsToCoord(10200));
    PositionE[5] := Point(MilsToCoord(7100), MilsToCoord(10100));
    PositionE[6] := Point(MilsToCoord(7100), MilsToCoord(10000));
    PositionE[7] := Point(MilsToCoord(7100), MilsToCoord(9900));
    PositionE[8] := Point(MilsToCoord(7100), MilsToCoord(9800));
    PositionE[9] := Point(MilsToCoord(7100), MilsToCoord(9700));
    PositionE[10] := Point(MilsToCoord(7100), MilsToCoord(9600));
    PositionE[11] := Point(MilsToCoord(7100), MilsToCoord(9500));
    PositionE[12] := Point(MilsToCoord(7100), MilsToCoord(9400));
    PositionE[13] := Point(MilsToCoord(7100), MilsToCoord(9300));
    PositionE[14] := Point(MilsToCoord(7100), MilsToCoord(9200));
    PositionE[15] := Point(MilsToCoord(7100), MilsToCoord(9100));
    PositionE[16] := Point(MilsToCoord(7100), MilsToCoord(9000));
    PositionE[17] := Point(MilsToCoord(7100), MilsToCoord(8900));
    PositionE[18] := Point(MilsToCoord(7100), MilsToCoord(8800));
    PositionE[19] := Point(MilsToCoord(7100), MilsToCoord(8700));
    PositionE[20] := Point(MilsToCoord(7100), MilsToCoord(8600));
    PositionE[21] := Point(MilsToCoord(7100), MilsToCoord(8500));
    //F
    PositionF[0] := Point(MilsToCoord(8600), MilsToCoord(10600));
    PositionF[1] := Point(MilsToCoord(8600), MilsToCoord(10500));
    PositionF[2] := Point(MilsToCoord(8600), MilsToCoord(10400));
    PositionF[3] := Point(MilsToCoord(8600), MilsToCoord(10300));
    PositionF[4] := Point(MilsToCoord(8600), MilsToCoord(10200));
    PositionF[5] := Point(MilsToCoord(8600), MilsToCoord(10100));
    PositionF[6] := Point(MilsToCoord(8600), MilsToCoord(10000));
    PositionF[7] := Point(MilsToCoord(8600), MilsToCoord(9900));
    PositionF[8] := Point(MilsToCoord(8600), MilsToCoord(9800));
    PositionF[9] := Point(MilsToCoord(8600), MilsToCoord(9700));
    PositionF[10] := Point(MilsToCoord(8600), MilsToCoord(9600));
    PositionF[11] := Point(MilsToCoord(8600), MilsToCoord(9500));
    PositionF[12] := Point(MilsToCoord(8600), MilsToCoord(9400));
    PositionF[13] := Point(MilsToCoord(8600), MilsToCoord(9300));
    PositionF[14] := Point(MilsToCoord(8600), MilsToCoord(9200));
    PositionF[15] := Point(MilsToCoord(8600), MilsToCoord(9100));
    PositionF[16] := Point(MilsToCoord(8600), MilsToCoord(9000));
    PositionF[17] := Point(MilsToCoord(8600), MilsToCoord(8900));
    PositionF[18] := Point(MilsToCoord(8600), MilsToCoord(8800));
    PositionF[19] := Point(MilsToCoord(8600), MilsToCoord(8700));
    PositionF[20] := Point(MilsToCoord(8600), MilsToCoord(8600));
    PositionF[21] := Point(MilsToCoord(8600), MilsToCoord(8500));

    //G
    PositionG[0] := Point(MilsToCoord(10100), MilsToCoord(10600));
    PositionG[1] := Point(MilsToCoord(10100), MilsToCoord(10500));
    PositionG[2] := Point(MilsToCoord(10100), MilsToCoord(10400));
    PositionG[3] := Point(MilsToCoord(10100), MilsToCoord(10300));
    PositionG[4] := Point(MilsToCoord(10100), MilsToCoord(10200));
    PositionG[5] := Point(MilsToCoord(10100), MilsToCoord(10100));
    PositionG[6] := Point(MilsToCoord(10100), MilsToCoord(10000));
    PositionG[7] := Point(MilsToCoord(10100), MilsToCoord(9900));
    PositionG[8] := Point(MilsToCoord(10100), MilsToCoord(9800));
    PositionG[9] := Point(MilsToCoord(10100), MilsToCoord(9700));
    PositionG[10] := Point(MilsToCoord(10100), MilsToCoord(9600));
    PositionG[11] := Point(MilsToCoord(10100), MilsToCoord(9500));
    PositionG[12] := Point(MilsToCoord(10100), MilsToCoord(9400));
    PositionG[13] := Point(MilsToCoord(10100), MilsToCoord(9300));
    PositionG[14] := Point(MilsToCoord(10100), MilsToCoord(9200));
    PositionG[15] := Point(MilsToCoord(10100), MilsToCoord(9100));
    PositionG[16] := Point(MilsToCoord(10100), MilsToCoord(9000));
    PositionG[17] := Point(MilsToCoord(10100), MilsToCoord(8900));
    PositionG[18] := Point(MilsToCoord(10100), MilsToCoord(8800));
    PositionG[19] := Point(MilsToCoord(10100), MilsToCoord(8700));
    PositionG[20] := Point(MilsToCoord(10100), MilsToCoord(8600));
    PositionG[21] := Point(MilsToCoord(10100), MilsToCoord(8500));

    //H
    PositionH[0] := Point(MilsToCoord(11800), MilsToCoord(10600));
    PositionH[1] := Point(MilsToCoord(11800), MilsToCoord(10500));
    PositionH[2] := Point(MilsToCoord(11800), MilsToCoord(10400));
    PositionH[3] := Point(MilsToCoord(11800), MilsToCoord(10300));
    PositionH[4] := Point(MilsToCoord(11800), MilsToCoord(10200));
    PositionH[5] := Point(MilsToCoord(11800), MilsToCoord(10100));
    PositionH[6] := Point(MilsToCoord(11800), MilsToCoord(10000));
    PositionH[7] := Point(MilsToCoord(11800), MilsToCoord(9900));
    PositionH[8] := Point(MilsToCoord(11800), MilsToCoord(9800));
    PositionH[9] := Point(MilsToCoord(11800), MilsToCoord(9700));
    PositionH[10] := Point(MilsToCoord(11800), MilsToCoord(9600));
    PositionH[11] := Point(MilsToCoord(11800), MilsToCoord(9500));
    PositionH[12] := Point(MilsToCoord(11800), MilsToCoord(9400));
    PositionH[13] := Point(MilsToCoord(11800), MilsToCoord(9300));
    PositionH[14] := Point(MilsToCoord(11800), MilsToCoord(9200));
    PositionH[15] := Point(MilsToCoord(11800), MilsToCoord(9100));
    PositionH[16] := Point(MilsToCoord(11800), MilsToCoord(9000));
    PositionH[17] := Point(MilsToCoord(11800), MilsToCoord(8900));
    PositionH[18] := Point(MilsToCoord(11800), MilsToCoord(8800));
    PositionH[19] := Point(MilsToCoord(11800), MilsToCoord(8700));
    PositionH[20] := Point(MilsToCoord(11800), MilsToCoord(8600));
    PositionH[21] := Point(MilsToCoord(11800), MilsToCoord(8500));

    //I
    PositionI[0] := Point(MilsToCoord(13300), MilsToCoord(10600));
    PositionI[1] := Point(MilsToCoord(13300), MilsToCoord(10500));
    PositionI[2] := Point(MilsToCoord(13300), MilsToCoord(10400));
    PositionI[3] := Point(MilsToCoord(13300), MilsToCoord(10300));
    PositionI[4] := Point(MilsToCoord(13300), MilsToCoord(10200));
    PositionI[5] := Point(MilsToCoord(13300), MilsToCoord(10100));
    PositionI[6] := Point(MilsToCoord(13300), MilsToCoord(10000));
    PositionI[7] := Point(MilsToCoord(13300), MilsToCoord(9900));
    PositionI[8] := Point(MilsToCoord(13300), MilsToCoord(9800));
    PositionI[9] := Point(MilsToCoord(13300), MilsToCoord(9700));
    PositionI[10] := Point(MilsToCoord(13300), MilsToCoord(9600));
    PositionI[11] := Point(MilsToCoord(13300), MilsToCoord(9500));
    PositionI[12] := Point(MilsToCoord(13300), MilsToCoord(9400));
    PositionI[13] := Point(MilsToCoord(13300), MilsToCoord(9300));
    PositionI[14] := Point(MilsToCoord(13300), MilsToCoord(9200));
    PositionI[15] := Point(MilsToCoord(13300), MilsToCoord(9100));
    PositionI[16] := Point(MilsToCoord(13300), MilsToCoord(9000));
    PositionI[17] := Point(MilsToCoord(13300), MilsToCoord(8900));
    PositionI[18] := Point(MilsToCoord(13300), MilsToCoord(8800));
    PositionI[19] := Point(MilsToCoord(13300), MilsToCoord(8700));
    PositionI[20] := Point(MilsToCoord(13300), MilsToCoord(8600));
    PositionI[21] := Point(MilsToCoord(13300), MilsToCoord(8500));

     // PositionJ 坐标
    PositionJ[0] := Point(MilsToCoord(800), MilsToCoord(8000));
    PositionJ[1] := Point(MilsToCoord(800), MilsToCoord(7900));
    PositionJ[2] := Point(MilsToCoord(800), MilsToCoord(7800));
    PositionJ[3] := Point(MilsToCoord(800), MilsToCoord(7700));
    PositionJ[4] := Point(MilsToCoord(800), MilsToCoord(7600));
    PositionJ[5] := Point(MilsToCoord(800), MilsToCoord(7500));
    PositionJ[6] := Point(MilsToCoord(800), MilsToCoord(7400));
    PositionJ[7] := Point(MilsToCoord(800), MilsToCoord(7300));
    PositionJ[8] := Point(MilsToCoord(800), MilsToCoord(7200));
    PositionJ[9] := Point(MilsToCoord(800), MilsToCoord(7100));
    PositionJ[10] := Point(MilsToCoord(800), MilsToCoord(7000));
    PositionJ[11] := Point(MilsToCoord(800), MilsToCoord(6900));
    PositionJ[12] := Point(MilsToCoord(800), MilsToCoord(6800));
    PositionJ[13] := Point(MilsToCoord(800), MilsToCoord(6700));
    PositionJ[14] := Point(MilsToCoord(800), MilsToCoord(6600));
    PositionJ[15] := Point(MilsToCoord(800), MilsToCoord(6500));
    PositionJ[16] := Point(MilsToCoord(800), MilsToCoord(6400));
    PositionJ[17] := Point(MilsToCoord(800), MilsToCoord(6300));
    PositionJ[18] := Point(MilsToCoord(800), MilsToCoord(6200));
    PositionJ[19] := Point(MilsToCoord(800), MilsToCoord(6100));
    PositionJ[20] := Point(MilsToCoord(800), MilsToCoord(6000));
    PositionJ[21] := Point(MilsToCoord(800), MilsToCoord(5900));

    // PositionK 坐标
    PositionK[0] := Point(MilsToCoord(2500), MilsToCoord(8000));
    PositionK[1] := Point(MilsToCoord(2500), MilsToCoord(7900));
    PositionK[2] := Point(MilsToCoord(2500), MilsToCoord(7800));
    PositionK[3] := Point(MilsToCoord(2500), MilsToCoord(7700));
    PositionK[4] := Point(MilsToCoord(2500), MilsToCoord(7600));
    PositionK[5] := Point(MilsToCoord(2500), MilsToCoord(7500));
    PositionK[6] := Point(MilsToCoord(2500), MilsToCoord(7400));
    PositionK[7] := Point(MilsToCoord(2500), MilsToCoord(7300));
    PositionK[8] := Point(MilsToCoord(2500), MilsToCoord(7200));
    PositionK[9] := Point(MilsToCoord(2500), MilsToCoord(7100));
    PositionK[10] := Point(MilsToCoord(2500), MilsToCoord(7000));
    PositionK[11] := Point(MilsToCoord(2500), MilsToCoord(6900));
    PositionK[12] := Point(MilsToCoord(2500), MilsToCoord(6800));
    PositionK[13] := Point(MilsToCoord(2500), MilsToCoord(6700));
    PositionK[14] := Point(MilsToCoord(2500), MilsToCoord(6600));
    PositionK[15] := Point(MilsToCoord(2500), MilsToCoord(6500));
    PositionK[16] := Point(MilsToCoord(2500), MilsToCoord(6400));
    PositionK[17] := Point(MilsToCoord(2500), MilsToCoord(6300));
    PositionK[18] := Point(MilsToCoord(2500), MilsToCoord(6200));
    PositionK[19] := Point(MilsToCoord(2500), MilsToCoord(6100));
    PositionK[20] := Point(MilsToCoord(2500), MilsToCoord(6000));
    PositionK[21] := Point(MilsToCoord(2500), MilsToCoord(5900));

    // PositionL 坐标
    PositionL[0] := Point(MilsToCoord(3900), MilsToCoord(8000));
    PositionL[1] := Point(MilsToCoord(3900), MilsToCoord(7900));
    PositionL[2] := Point(MilsToCoord(3900), MilsToCoord(7800));
    PositionL[3] := Point(MilsToCoord(3900), MilsToCoord(7700));
    PositionL[4] := Point(MilsToCoord(3900), MilsToCoord(7600));
    PositionL[5] := Point(MilsToCoord(3900), MilsToCoord(7500));
    PositionL[6] := Point(MilsToCoord(3900), MilsToCoord(7400));
    PositionL[7] := Point(MilsToCoord(3900), MilsToCoord(7300));
    PositionL[8] := Point(MilsToCoord(3900), MilsToCoord(7200));
    PositionL[9] := Point(MilsToCoord(3900), MilsToCoord(7100));
    PositionL[10] := Point(MilsToCoord(3900), MilsToCoord(7000));
    PositionL[11] := Point(MilsToCoord(3900), MilsToCoord(6900));
    PositionL[12] := Point(MilsToCoord(3900), MilsToCoord(6800));
    PositionL[13] := Point(MilsToCoord(3900), MilsToCoord(6700));
    PositionL[14] := Point(MilsToCoord(3900), MilsToCoord(6600));
    PositionL[15] := Point(MilsToCoord(3900), MilsToCoord(6500));
    PositionL[16] := Point(MilsToCoord(3900), MilsToCoord(6400));
    PositionL[17] := Point(MilsToCoord(3900), MilsToCoord(6300));
    PositionL[18] := Point(MilsToCoord(3900), MilsToCoord(6200));
    PositionL[19] := Point(MilsToCoord(3900), MilsToCoord(6100));
    PositionL[20] := Point(MilsToCoord(3900), MilsToCoord(6000));
    PositionL[21] := Point(MilsToCoord(3900), MilsToCoord(5900));

    // PositionM 坐标
    PositionM[0] := Point(MilsToCoord(5600), MilsToCoord(8000));
    PositionM[1] := Point(MilsToCoord(5600), MilsToCoord(7900));
    PositionM[2] := Point(MilsToCoord(5600), MilsToCoord(7800));
    PositionM[3] := Point(MilsToCoord(5600), MilsToCoord(7700));
    PositionM[4] := Point(MilsToCoord(5600), MilsToCoord(7600));
    PositionM[5] := Point(MilsToCoord(5600), MilsToCoord(7500));
    PositionM[6] := Point(MilsToCoord(5600), MilsToCoord(7400));
    PositionM[7] := Point(MilsToCoord(5600), MilsToCoord(7300));
    PositionM[8] := Point(MilsToCoord(5600), MilsToCoord(7200));
    PositionM[9] := Point(MilsToCoord(5600), MilsToCoord(7100));
    PositionM[10] := Point(MilsToCoord(5600), MilsToCoord(7000));
    PositionM[11] := Point(MilsToCoord(5600), MilsToCoord(6900));
    PositionM[12] := Point(MilsToCoord(5600), MilsToCoord(6800));
    PositionM[13] := Point(MilsToCoord(5600), MilsToCoord(6700));
    PositionM[14] := Point(MilsToCoord(5600), MilsToCoord(6600));
    PositionM[15] := Point(MilsToCoord(5600), MilsToCoord(6500));
    PositionM[16] := Point(MilsToCoord(5600), MilsToCoord(6400));
    PositionM[17] := Point(MilsToCoord(5600), MilsToCoord(6300));
    PositionM[18] := Point(MilsToCoord(5600), MilsToCoord(6200));
    PositionM[19] := Point(MilsToCoord(5600), MilsToCoord(6100));
    PositionM[20] := Point(MilsToCoord(5600), MilsToCoord(6000));
    PositionM[21] := Point(MilsToCoord(5600), MilsToCoord(5900));

    // PositionN 坐标
    PositionN[0] := Point(MilsToCoord(7100), MilsToCoord(8000));
    PositionN[1] := Point(MilsToCoord(7100), MilsToCoord(7900));
    PositionN[2] := Point(MilsToCoord(7100), MilsToCoord(7800));
    PositionN[3] := Point(MilsToCoord(7100), MilsToCoord(7700));
    PositionN[4] := Point(MilsToCoord(7100), MilsToCoord(7600));
    PositionN[5] := Point(MilsToCoord(7100), MilsToCoord(7500));
    PositionN[6] := Point(MilsToCoord(7100), MilsToCoord(7400));
    PositionN[7] := Point(MilsToCoord(7100), MilsToCoord(7300));
    PositionN[8] := Point(MilsToCoord(7100), MilsToCoord(7200));
    PositionN[9] := Point(MilsToCoord(7100), MilsToCoord(7100));
    PositionN[10] := Point(MilsToCoord(7100), MilsToCoord(7000));
    PositionN[11] := Point(MilsToCoord(7100), MilsToCoord(6900));
    PositionN[12] := Point(MilsToCoord(7100), MilsToCoord(6800));
    PositionN[13] := Point(MilsToCoord(7100), MilsToCoord(6700));
    PositionN[14] := Point(MilsToCoord(7100), MilsToCoord(6600));
    PositionN[15] := Point(MilsToCoord(7100), MilsToCoord(6500));
    PositionN[16] := Point(MilsToCoord(7100), MilsToCoord(6400));
    PositionN[17] := Point(MilsToCoord(7100), MilsToCoord(6300));
    PositionN[18] := Point(MilsToCoord(7100), MilsToCoord(6200));
    PositionN[19] := Point(MilsToCoord(7100), MilsToCoord(6100));
    PositionN[20] := Point(MilsToCoord(7100), MilsToCoord(6000));
    PositionN[21] := Point(MilsToCoord(7100), MilsToCoord(5900));

    // PositionO 坐标
    PositionO[0] := Point(MilsToCoord(8600), MilsToCoord(8000));
    PositionO[1] := Point(MilsToCoord(8600), MilsToCoord(7900));
    PositionO[2] := Point(MilsToCoord(8600), MilsToCoord(7800));
    PositionO[3] := Point(MilsToCoord(8600), MilsToCoord(7700));
    PositionO[4] := Point(MilsToCoord(8600), MilsToCoord(7600));
    PositionO[5] := Point(MilsToCoord(8600), MilsToCoord(7500));
    PositionO[6] := Point(MilsToCoord(8600), MilsToCoord(7400));
    PositionO[7] := Point(MilsToCoord(8600), MilsToCoord(7300));
    PositionO[8] := Point(MilsToCoord(8600), MilsToCoord(7200));
    PositionO[9] := Point(MilsToCoord(8600), MilsToCoord(7100));
    PositionO[10] := Point(MilsToCoord(8600), MilsToCoord(7000));
    PositionO[11] := Point(MilsToCoord(8600), MilsToCoord(6900));
    PositionO[12] := Point(MilsToCoord(8600), MilsToCoord(6800));
    PositionO[13] := Point(MilsToCoord(8600), MilsToCoord(6700));
    PositionO[14] := Point(MilsToCoord(8600), MilsToCoord(6600));
    PositionO[15] := Point(MilsToCoord(8600), MilsToCoord(6500));
    PositionO[16] := Point(MilsToCoord(8600), MilsToCoord(6400));
    PositionO[17] := Point(MilsToCoord(8600), MilsToCoord(6300));
    PositionO[18] := Point(MilsToCoord(8600), MilsToCoord(6200));
    PositionO[19] := Point(MilsToCoord(8600), MilsToCoord(6100));
    PositionO[20] := Point(MilsToCoord(8600), MilsToCoord(6000));
    PositionO[21] := Point(MilsToCoord(8600), MilsToCoord(5900));

    // PositionP 坐标
    PositionP[0] := Point(MilsToCoord(10100), MilsToCoord(8000));
    PositionP[1] := Point(MilsToCoord(10100), MilsToCoord(7900));
    PositionP[2] := Point(MilsToCoord(10100), MilsToCoord(7800));
    PositionP[3] := Point(MilsToCoord(10100), MilsToCoord(7700));
    PositionP[4] := Point(MilsToCoord(10100), MilsToCoord(7600));
    PositionP[5] := Point(MilsToCoord(10100), MilsToCoord(7500));
    PositionP[6] := Point(MilsToCoord(10100), MilsToCoord(7400));
    PositionP[7] := Point(MilsToCoord(10100), MilsToCoord(7300));
    PositionP[8] := Point(MilsToCoord(10100), MilsToCoord(7200));
    PositionP[9] := Point(MilsToCoord(10100), MilsToCoord(7100));
    PositionP[10] := Point(MilsToCoord(10100), MilsToCoord(7000));
    PositionP[11] := Point(MilsToCoord(10100), MilsToCoord(6900));
    PositionP[12] := Point(MilsToCoord(10100), MilsToCoord(6800));
    PositionP[13] := Point(MilsToCoord(10100), MilsToCoord(6700));
    PositionP[14] := Point(MilsToCoord(10100), MilsToCoord(6600));
    PositionP[15] := Point(MilsToCoord(10100), MilsToCoord(6500));
    PositionP[16] := Point(MilsToCoord(10100), MilsToCoord(6400));
    PositionP[17] := Point(MilsToCoord(10100), MilsToCoord(6300));
    PositionP[18] := Point(MilsToCoord(10100), MilsToCoord(6200));
    PositionP[19] := Point(MilsToCoord(10100), MilsToCoord(6100));
    PositionP[20] := Point(MilsToCoord(10100), MilsToCoord(6000));
    PositionP[21] := Point(MilsToCoord(10100), MilsToCoord(5900));

    // PositionQ 坐标
    PositionQ[0] := Point(MilsToCoord(11800), MilsToCoord(8000));
    PositionQ[1] := Point(MilsToCoord(11800), MilsToCoord(7900));
    PositionQ[2] := Point(MilsToCoord(11800), MilsToCoord(7800));
    PositionQ[3] := Point(MilsToCoord(11800), MilsToCoord(7700));
    PositionQ[4] := Point(MilsToCoord(11800), MilsToCoord(7600));
    PositionQ[5] := Point(MilsToCoord(11800), MilsToCoord(7500));
    PositionQ[6] := Point(MilsToCoord(11800), MilsToCoord(7400));
    PositionQ[7] := Point(MilsToCoord(11800), MilsToCoord(7300));
    PositionQ[8] := Point(MilsToCoord(11800), MilsToCoord(7200));
    PositionQ[9] := Point(MilsToCoord(11800), MilsToCoord(7100));
    PositionQ[10] := Point(MilsToCoord(11800), MilsToCoord(7000));
    PositionQ[11] := Point(MilsToCoord(11800), MilsToCoord(6900));
    PositionQ[12] := Point(MilsToCoord(11800), MilsToCoord(6800));
    PositionQ[13] := Point(MilsToCoord(11800), MilsToCoord(6700));
    PositionQ[14] := Point(MilsToCoord(11800), MilsToCoord(6600));
    PositionQ[15] := Point(MilsToCoord(11800), MilsToCoord(6500));
    PositionQ[16] := Point(MilsToCoord(11800), MilsToCoord(6400));
    PositionQ[17] := Point(MilsToCoord(11800), MilsToCoord(6300));
    PositionQ[18] := Point(MilsToCoord(11800), MilsToCoord(6200));
    PositionQ[19] := Point(MilsToCoord(11800), MilsToCoord(6100));
    PositionQ[20] := Point(MilsToCoord(11800), MilsToCoord(6000));
    PositionQ[21] := Point(MilsToCoord(11800), MilsToCoord(5900));




    // Get current SchDoc
    SchDoc := SchServer.GetCurrentSchDocument;
    if SchDoc = nil then
    begin
        ShowMessage('Please input a SchDOc Path');
        Exit;
    end;

    // Open CSV File
    AssignFile(CSVFile, CSVFilePath);
    Reset(CSVFile);

    try
        // Skip the tile
        Readln(CSVFile, Line);

        // put data into 22 Positions
        i := 0;
        while (not EOF(CSVFile)) and (i < 22) do
        begin
            Readln(CSVFile, Line);
    // Skip empty or whitespace-only lines
    Line := Trim(Line);
    if (Line = '') or (Length(Line) = 0) then Continue;

    // Skip lines that don't have enough fields
    if Pos(',', Line) = 0 then Continue;

            Fields := TStringList.Create;
            Fields.Delimiter := ',';
            Fields.StrictDelimiter := True;
            Fields.DelimitedText := Line;

            // Pick the data of Col
            if Fields.Count > 1 then
            begin
                PortNamesA[i] := Fields[0]; //
                PortNamesB[i] := Fields[1]; //
                PortNamesC[i] := Fields[2];
                PortNamesD[i] := Fields[3];
                PortNamesE[i] := Fields[4];
                PortNamesF[i] := Fields[5];
                PortNamesG[i] := Fields[6];
                PortNamesH[i] := Fields[7];
                PortNamesI[i] := Fields[8];
                PortNamesJ[i] := Fields[9];
                PortNamesK[i] := Fields[10];
                PortNamesL[i] := Fields[11];
                PortNamesM[i] := Fields[12];
                PortNamesN[i] := Fields[13];
                PortNamesO[i] := Fields[14];
                PortNamesP[i] := Fields[15];
                PortNamesQ[i] := Fields[16];


            end;

            Fields.Free;
            Inc(i);
        end;



        // 在 PositionA 位置生成第一列 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesA[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionA[i];
                    PowerPort.Text := PortNamesA[i];
                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionB 位置生成第二列 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesB[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionB[i];
                    PowerPort.Text := PortNamesB[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionC 位置生成第二列 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesC[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionC[i];
                    PowerPort.Text := PortNamesC[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionB 位置生成第二列 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesD[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionD[i];
                    PowerPort.Text := PortNamesD[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionB 位置生成第二列 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesE[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionE[i];
                    PowerPort.Text := PortNamesE[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;


        // 在 PositionF 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesF[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionF[i];
                    PowerPort.Text := PortNamesF[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionG 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesG[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionG[i];
                    PowerPort.Text := PortNamesG[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionH 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesH[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionH[i];
                    PowerPort.Text := PortNamesH[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionI 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesI[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionI[i];
                    PowerPort.Text := PortNamesI[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionJ 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesJ[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionJ[i];
                    PowerPort.Text := PortNamesJ[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionK 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesK[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionK[i];
                    PowerPort.Text := PortNamesK[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionL 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesL[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionL[i];
                    PowerPort.Text := PortNamesL[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionM 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesM[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionM[i];
                    PowerPort.Text := PortNamesM[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionN 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesN[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionN[i];
                    PowerPort.Text := PortNamesN[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionO 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesO[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionO[i];
                    PowerPort.Text := PortNamesO[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionP 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesP[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionP[i];
                    PowerPort.Text := PortNamesP[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);

                end;
            end;
        end;

        // 在 PositionQ 位置生成 Power Ports
        for i := 0 to 21 do
        begin
            if PortNamesQ[i] <> '' then  // 检查是否为空
            begin
                PowerPort := SchServer.SchObjectFactory(ePowerObject, eCreate_Default);
                if PowerPort <> nil then
                begin
                    PowerPort.Location := PositionQ[i];
                    PowerPort.Text := PortNamesQ[i];

                    SchDoc.RegisterSchObjectInContainer(PowerPort);
                    
                end;
            end;
        end;


    finally
        // 关闭 CSV 文件
        CloseFile(CSVFile);
    end;

    ShowMessage('Power ports already imported into SchDoc successfully');
end;


// Change Path here
procedure ImportPowerPorts_Octant1;
begin
    ImportPowerPortsFromCSV('C:\Users\63263\Desktop\OCT1_1.csv');
end;

procedure ImportPowerPorts_Octant2;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\OCT2.csv');
end;

procedure ImportPowerPorts_Octant3;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\OCT3.csv');
end;

procedure ImportPowerPorts_Octant4;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\OCT4.csv');
end;

procedure ImportPowerPorts_Octant5;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\OCT5.csv');
end;

procedure ImportPowerPorts_Octant6;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\OCT6.csv');
end;

procedure ImportPowerPorts_Octant7;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\OCT7.csv');
end;

procedure ImportPowerPorts_Octant8;
begin
    ImportPowerPortsFromCSV('D:\Acco_脚本\AD8300自动脚本\OCT8.csv');
end;


