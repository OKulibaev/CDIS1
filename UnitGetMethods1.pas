unit UnitGetMethods1;

interface

uses
  Windows, Messages, System.SysUtils, System.UITypes, Vcl.Graphics, Variants,
  Controls, StdCtrls, Classes, Forms, DateUtils, Dialogs, Math, DB, MemDS,
  DBAccess, MSAccess, UnitMyClasses1, dxTileControl, dxRatingControl,
  cxDateUtils, cxGridChartView, cxGridDBBandedTableView, cxGridDBChartView,
  cxGridBandedTableView, cxGridServerModeBandedTableView, cxGridDBCardView,
  cxGridTableView, cxTL, cxClasses, cxStyles, cxGraphics, cxEdit, cxImage,
  cxCalendar, cxCalc, cxCheckBox, cxImageComboBox, cxTimeEdit,
  cxDBLookupComboBox, cxMaskEdit, cxGroupBox;

  function GetAllTreeChildRecords1(aTreeLV2SA1: TcxTreeList; aRecord_IDLSV16: string): string;
  function GetColourFromColourName1(aColour_NameTSV59G: string): TColor;
  function GetDateTimeStrValueFromQuery1(stringTempSQL593G, stringValue593, aFormatSV595G: string): string;
  function GetDateTimeValueFromQuery1(stringTempSQL594G, stringValue594: string): TDateTime;
  function GetDetail_IDFromANamedField1(aTable_IDLSV3H, aText_Field_ValueSV3H, aField_NameLSV3H: string): string;
  function GetDetail_IDFromAnIndexTitle1(aTable_IDLSV3H, aText_Field_ValueSV3H, aField_NameLSV3H: string): string;
  function GetEncrypted_SQL_Statement1(aValueLSV56: string): string;
  function GetFileInfoCreateDate1(aStrPath1: string): string;
  function GetFileInfoSize1(aStrPath1: string): string;
  function GetFloatValueFromQuery1(stringTempSQL592G, stringValue592: string): Double;
  function GetHeightOfTaskBar1: integer;
  function GetIntegerValueFromQuery1(stringTempSQL591G, stringValue591: string): integer;
  function GetLatitudeString1(aLatitude1LSV3D: string): string;
  function GetLongitudeString1(aLongitude1LSV3D: string): string;
  function GetMaxNewlyCreatedID1(aCurrentDataGridLV317J: TaCurrentDataTableStructure): string;
  function GetMaxVoyDate1(aDate_IDLSV23TR4, aUnit_IDLSV23TR4: string): TDateTime;
  function GetMaxVoyDoubleDataFromUnitID1(aField_NameLSV23TR4, aUnit_IDLSV23TR4: string): Double;
  function GetNoOfConnectedRecords1(aDaughter_IDIDLSV87F, aParent_IDLSV87F, aParent_IndexLSV87F: string): integer;
  function GetNoOfMessagesWithThisID1(aMessage_ID1LSV4: string): integer;
  function GetNoOfNewNotifications1: integer;
  function GetNoOfGPSBetweenDates1(dtStart24, dtFinish24: TDateTime; stringMap_Ship_ID_List24: string): integer;
  function GetNoOfPosRepsBetweenDates1(dtStart25, dtFinish25: TDateTime; stringMap_Ship_ID_List25: string = ''): integer;
  function GetNoOfUserActiveTasks1: integer;
  function GetPerson_IDFromLogin_ID(aLogin_IDLVS66: string): string;
  function GetPerson_IDFromNameSurname1(stringPerson_Name1: string): string;
  function GetPersonLoginDateTimeFromPerson_ID(Person_IDLVS66: string): string;
  function GetPersonNameSurnameFromPerson_ID1(aPerson_IDLVS66: string): string;
  function GetRecord_Type_Detail1(aDetail_Field_NameIDLSV843, aRecord_Type_IDLSV843: string): string;
  function GetRecord_Type_Detail2(aDetail_Field_NameIDLSV843, aRecord_Type_Index_IDLSV843: string): string;
  function GetRecord_Type_Colour1(stringRecord_Type_ID1: string): integer;
  function GetColourNumberFromColour_ID1(stringColour_ID1: string): integer;
  function GetShip_ColourFromShip_ID1(stringShip_ID1: string): integer;
  function GetShip_IDFromShipData_IMO_Number1(aShipDataLV34D: TaShipData): string;
  function GetStringValueFromAnyValue1(aField_Name1LV41D, aField_Name2LV41D, aTable_NameLV41D, aValueLV41D: string): string;
  function GetStringValueFromFloat1(aTempQueryLV4D: TMSQuery; stringFname4D: string; intPrecision1: integer): string;
  function GetStringValueFromQuery1(stringTempSQL13R5, stringRESN1: string): string;
  function GetTilesButtonsWidth1(aTileControlLV34 : TdxTileControl): Integer;
  function GetTimeZoneDoubleFromString1(aTime_Zone_StringLSV34: string): Double;
  function GetTopPanelPosition1L1(aPanel: TcxGroupBox): integer;
  function GetUpdateStringWithIntegerFromCurrentDataDetailFromIndex1(aCurrentDetailsLV88LPZ: TaRecordTypeStrings; aFied_NameLSV88LPZ, aField_ValueLSV88LPZ: string): string;
  function GetUpdateStringWithStringFromCurrentDataDetailFromIndex1(aCurrentDetailsLV88LPZ: TaRecordTypeStrings; aFied_NameLSV88LPZ, aField_ValueLSV88LPZ: string): string;
  function GetVoyage_IDFromShipData_Ship_IDAndVoyageNumber1(aShipDataLV34D: TaShipData): string;

implementation

uses
  UnitMain,
  UnitDM1,
  UnitMiscMethods1,

  UnitMyConstants1;


function GetStringValueFromQuery1(stringTempSQL13R5, stringRESN1: string): string;
var
  aQueryTM4: TMSQuery;
begin
  try
    aQueryTM4 := TMSQuery.Create(nil);
    try
      Result := '';
      try
        if stringTempSQL13R5 <> '' then
        begin
          if FormDM1.Database_Connection_SDAC.Connected = True then
          begin
            aQueryTM4.Connection := FormDM1.Database_Connection_SDAC;
            aQueryTM4.Close;
            aQueryTM4.SQL.Clear;
            aQueryTM4.SQL.Add(stringTempSQL13R5);
            aQueryTM4.Open;
            if aQueryTM4.RecordCount > 0 then
              Result := Trim(aQueryTM4.FieldByName(stringRESN1).AsString);
            aQueryTM4.Close;
          end
          else
            SaveMsgToAFile1(constLogFolderName1, 'No Database Connection!', 'GetStringValueFromQuery1.NoDBConnectionError');
        end
        else
          SaveMsgToAFile1(constLogFolderName1, 'No SQL Text!', 'GetStringValueFromQuery1.NoSQLError');
      except
        Exit;
      end;
    except
      on E1: Exception do
      begin
        Result := '';
        GlobSystemData1.aLogError1(E1, 'GetStringValueFromQuery1', stringTempSQL13R5 + ', ' + stringRESN1);
        Exit;
      end;
    end;
  finally
    FreeAndNil(aQueryTM4);
  end;
end;

function GetIntegerValueFromQuery1(stringTempSQL591G, stringValue591: string): integer;
var
  aQueryTM5: TMSQuery;
begin
  try
    aQueryTM5 := TMSQuery.Create(nil);
    try
      Result := 0;
      try
        if (Trim(stringTempSQL591G) <> '') AND (UpperCase(Trim(stringTempSQL591G)) <> 'N/A') then
        begin
          if FormDM1.Database_Connection_SDAC.Connected = True then
          begin
            aQueryTM5.Connection := FormDM1.Database_Connection_SDAC;
            aQueryTM5.Close;
            aQueryTM5.SQL.Clear;
            aQueryTM5.SQL.Add(stringTempSQL591G);
            aQueryTM5.Open;
            if aQueryTM5.RecordCount > 0 then
              Result := aQueryTM5.FieldByName(stringValue591).AsInteger;
            aQueryTM5.Close;
          end
          else
            SaveMsgToAFile1(constLogFolderName1, 'No Database Connection!', 'GetIntegerValueFromQuery1.NoDBConnectionError');
        end
        else
          SaveMsgToAFile1(constLogFolderName1, 'No SQL Text!', 'GetIntegerValueFromQuery1.NoSQLError');
      except
        Exit;
      end;
    except
      on E1: Exception do
      begin
        Result := 0;
        GlobSystemData1.aLogError1(E1, 'GetIntegerValueFromQuery1', stringTempSQL591G + ', ' + stringValue591);
        Exit;
      end;
    end;
  finally
    FreeAndNil(aQueryTM5);
  end;
end;

function GetFloatValueFromQuery1(stringTempSQL592G, stringValue592: string): Double;
var
  aQueryTM6: TMSQuery;
begin
  try
    aQueryTM6 := TMSQuery.Create(nil);
    try
      Result := 0;
      try
        if (Trim(stringTempSQL592G) <> '') AND (UpperCase(Trim(stringTempSQL592G)) <> 'N/A') then
        begin
          if FormDM1.Database_Connection_SDAC.Connected = True then
          begin
            aQueryTM6.Connection := FormDM1.Database_Connection_SDAC;
            aQueryTM6.Close;
            aQueryTM6.SQL.Clear;
            aQueryTM6.SQL.Add(stringTempSQL592G);
            aQueryTM6.Open;
            if aQueryTM6.RecordCount > 0 then
              Result := aQueryTM6.FieldByName(stringValue592).AsFloat;
            aQueryTM6.Close;
          end
          else
            SaveMsgToAFile1(constLogFolderName1, 'No Database Connection!', 'GetFloatValueFromQuery1.NoDBConnectionError');
        end
        else
          SaveMsgToAFile1(constLogFolderName1, 'No SQL Text!', 'GetFloatValueFromQuery1.NoSQLError');
      except
        Exit;
      end;
    except
      on E1: Exception do
      begin
        Result := 0;
        GlobSystemData1.aLogError1(E1, 'GetFloatValueFromQuery1', stringTempSQL592G + ', ' + stringValue592);
        Exit;
      end;
    end;
  finally
    FreeAndNil(aQueryTM6);
  end;
end;

function GetDateTimeStrValueFromQuery1(stringTempSQL593G, stringValue593, aFormatSV595G: string): string;
var
  aQueryTM7: TMSQuery;
begin
  try
    aQueryTM7 := TMSQuery.Create(nil);
    try
      Result := '';
      try
        if (Trim(stringTempSQL593G) <> '') AND (UpperCase(Trim(stringTempSQL593G)) <> 'N/A') then
        begin
          if FormDM1.Database_Connection_SDAC.Connected = True then
          begin
            aQueryTM7.Connection := FormDM1.Database_Connection_SDAC;
            aQueryTM7.Close;
            aQueryTM7.SQL.Clear;
            aQueryTM7.SQL.Add(stringTempSQL593G);
            aQueryTM7.Open;
            if aQueryTM7.RecordCount > 0 then
            begin
              if aQueryTM7.FieldByName(stringValue593).IsNull = False then
                Result := FormatDateTime(aFormatSV595G, aQueryTM7.FieldByName(stringValue593).AsDateTime);
            end;
            aQueryTM7.Close;
          end
          else
            SaveMsgToAFile1(constLogFolderName1, 'No Database Connection!', 'GetDateTimeStrValueFromQuery1.NoDBConnectionError');
        end
        else
          SaveMsgToAFile1(constLogFolderName1, 'No SQL Text!', 'GetDateTimeStrValueFromQuery1.NoSQLError');
      except
        Exit;
      end;
    except
      on E1: Exception do
      begin
        Result := '';
        GlobSystemData1.aLogError1(E1, 'GetDateTimeStrValueFromQuery1', stringTempSQL593G + ', ' + stringValue593);
        Exit;
      end;
    end;
  finally
    FreeAndNil(aQueryTM7);
  end;
end;

function GetDateTimeValueFromQuery1(stringTempSQL594G, stringValue594: string): TDateTime;
var
  aQueryTM8: TMSQuery;
begin
  try
    aQueryTM8 := TMSQuery.Create(nil);
    try
      Result := NullDate;
      try
        if (Trim(stringTempSQL594G) <> '') AND (UpperCase(Trim(stringTempSQL594G)) <> 'N/A') then
        begin
          if FormDM1.Database_Connection_SDAC.Connected = True then
          begin
            aQueryTM8.Connection := FormDM1.Database_Connection_SDAC;
            aQueryTM8.Close;
            aQueryTM8.SQL.Clear;
            aQueryTM8.SQL.Add(stringTempSQL594G);
            aQueryTM8.Open;
            if aQueryTM8.RecordCount > 0 then
            begin
              if aQueryTM8.FieldByName(stringValue594).IsNull = False then Result := aQueryTM8.FieldByName(stringValue594).AsDateTime;
            end;
            aQueryTM8.Close;
          end
          else
            SaveMsgToAFile1(constLogFolderName1, 'No Database Connection!', 'GetDateTimeValueFromQuery1.NoDBConnectionError');
        end
        else
          SaveMsgToAFile1(constLogFolderName1, 'No SQL Text!', 'GetDateTimeValueFromQuery1.NoSQLError');
      except
        Exit;
      end;
    except
      on E1: Exception do
      begin
        Result := NullDate;
        GlobSystemData1.aLogError1(E1, 'GetDateTimeValueFromQuery1', stringTempSQL594G + ', ' + stringValue594);
        Exit;
      end;
    end;
  finally
    FreeAndNil(aQueryTM8);
  end;
end;

function GetColourFromColourName1(aColour_NameTSV59G: string): TColor;
begin
  try
    Result := $001D1D1D;
    try
      if UpperCase(aColour_NameTSV59G) = 'WHITE' then Result:= clCream else
       if UpperCase(aColour_NameTSV59G) = 'DARK ORANGE' then Result:= $002C53DA else
        if UpperCase(aColour_NameTSV59G) = 'DARK PURPLE' then Result:= $20BA3C60 else
         if UpperCase(aColour_NameTSV59G) = 'DARK RED' then Result:= $00471DB9 else
          if UpperCase(aColour_NameTSV59G) = 'GREEN' then Result:= $2000A300 else
           if UpperCase(aColour_NameTSV59G) = 'LIGHT GREEN' then Result:= $0033B499 else
            if UpperCase(aColour_NameTSV59G) = 'LIGHT ORANGE' then Result:= $201AA2E3 else
             if UpperCase(aColour_NameTSV59G) = 'LIGHT PURPLE' then Result:= $00A7009F else
              if UpperCase(aColour_NameTSV59G) = 'NAVY' then Result:= $0097572B else
               if UpperCase(aColour_NameTSV59G) = 'ORANGE' then Result:= $200068FA else
                if UpperCase(aColour_NameTSV59G) = 'PINK' then Result:= $209700FF else
                 if UpperCase(aColour_NameTSV59G) = 'RED' then Result:= $201111EE else
                  if UpperCase(aColour_NameTSV59G) = 'SKY BLUE' then Result:= $00FFF4EF else
                   if UpperCase(aColour_NameTSV59G) = 'TEAL' then Result:= $00A9AB00 else
                    if UpperCase(aColour_NameTSV59G) = 'VIOLET' then Result:= $0078387E else
                     if UpperCase(aColour_NameTSV59G) = 'YELLOW' then Result:= $000DC4FF else
                      if UpperCase(aColour_NameTSV59G) = 'BLUE' then Result:= $00EF892D else
                       if UpperCase(aColour_NameTSV59G) = 'DARK GREEN' then Result:= $0045711E;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := $001D1D1D;
      GlobSystemData1.aLogError1(E1, 'GetColourFromColourName1');
      Exit;
    end;
  end;
end;

function GetAllTreeChildRecords1(aTreeLV2SA1: TcxTreeList; aRecord_IDLSV16: string): string;
var
  aTreeListNode1, aTreeListNode2, aTreeListNode3, aTreeListNode4, aTreeListNode5: TcxTreeListNode;
  aRecord_Name1, aParent_Name1, aParent_Name2, aParent_Name3, aParent_Name4: string;
  aResult_StringLSV16: string;

  function aRecSFG(aParent_NodeLV1: TcxTreeListNode): TcxTreeListNode;
  var
    aShtring_Value1SFDF: string;
  begin
    try
      Result := nil;
      try
        aShtring_Value1SFDF := '';
        Result := aParent_NodeLV1.GetNext;
        aShtring_Value1SFDF := Trim(Result.Values[1]);
        if aShtring_Value1SFDF <> '' then
        begin
          if aResult_StringLSV16 <> '' then
            aResult_StringLSV16 := aResult_StringLSV16 + ',' + aShtring_Value1SFDF
          else
            aResult_StringLSV16 := aShtring_Value1SFDF;
        end;
        if Result <> aTreeListNode5 then
          Result := aRecSFG(Result);
      except
        Exit;
      end;
    except
      on E1: Exception do
      begin
        Result := nil;
        GlobSystemData1.aLogError1(E1, 'GetAllTreeChildRecords1.aRecSFG');
        Exit;
      end;
    end;
  end;

begin
  try
    Result := '';
    try
      aResult_StringLSV16 := '';
      aRecord_Name1 := '';
      aParent_Name1 := '';
      aParent_Name2 := '';
      aParent_Name3 := '';
      aParent_Name4 := '';
      aTreeListNode1 := aTreeLV2SA1.FindNodeByText(aRecord_IDLSV16, aTreeLV2SA1.Columns[1], nil, False, True, True, tlfmExact);
      if aTreeListNode1 <> nil then
      begin
        if aTreeListNode1.HasChildren = True then
        begin
          aTreeListNode2 := aTreeListNode1.GetLastChild;
          if aTreeListNode2.GetLastChild <> nil then
            aTreeListNode3 := aTreeListNode2.GetLastChild
          else
            aTreeListNode3 := aTreeListNode2;
          if aTreeListNode3.GetLastChild <> nil then
            aTreeListNode4 := aTreeListNode3.GetLastChild
          else
            aTreeListNode4 := aTreeListNode3;
          if aTreeListNode4.GetLastChild <> nil then
            aTreeListNode5 := aTreeListNode4.GetLastChild
          else
            aTreeListNode5 := aTreeListNode4;
          aRecSFG(aTreeListNode1);
        end;
        Result := aResult_StringLSV16;
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetAllTreeChildRecords1');
      Exit;
    end;
  end;
end;

function GetLatitudeString1(aLatitude1LSV3D: string): string;
var
  aLatitude1LSV3D1, aLatitude1LSV3D2, aLATSignLSV3D: string;
begin
  try
    Result := '0';
    try
      aLatitude1LSV3D1 := '';
      aLatitude1LSV3D2 := '';
      aLATSignLSV3D := '';
      aLatitude1LSV3D1 := System.Copy(aLatitude1LSV3D, 1, System.Pos('-', aLatitude1LSV3D) - 1);
      aLatitude1LSV3D2 := System.Copy(aLatitude1LSV3D, System.Pos('-', aLatitude1LSV3D) + 1, 2);
      if (aLatitude1LSV3D2 <> '') AND (aLatitude1LSV3D2 <> '00') then
        aLatitude1LSV3D2 := IntToStr(Round(StrToInt(aLatitude1LSV3D2) / 60 * 100));
      if Length(aLatitude1LSV3D2) = 1 then
        aLatitude1LSV3D2 := '0' + aLatitude1LSV3D2;
      if System.Pos('S', UpperCase(aLatitude1LSV3D)) > 0 then
        aLATSignLSV3D := '-'
      else
        aLATSignLSV3D := '';
      Result := aLATSignLSV3D + aLatitude1LSV3D1 + GlobLocaleFormats1.DecimalSeparator + aLatitude1LSV3D2;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '0';
      GlobSystemData1.aLogError1(E1, 'GetLatitudeString1');
      Exit;
    end;
  end;
end;

function GetLongitudeString1(aLongitude1LSV3D: string): string;
var
  aLongitude1LSV3D1, aLongitude1LSV3D2, aLONSignLSV3D: string;
begin
  try
    Result := '0';
    try
      aLongitude1LSV3D1 := '';
      aLongitude1LSV3D2 := '';
      aLONSignLSV3D := '';
      aLongitude1LSV3D1 := System.Copy(aLongitude1LSV3D, 1, System.Pos('-', aLongitude1LSV3D) - 1);
      aLongitude1LSV3D2 := System.Copy(aLongitude1LSV3D, System.Pos('-', aLongitude1LSV3D) + 1, 2);
      if (aLongitude1LSV3D2 <> '') AND (aLongitude1LSV3D2 <> '00') then
        aLongitude1LSV3D2 := IntToStr(Round(StrToInt(aLongitude1LSV3D2) / 60 * 100));
      if Length(aLongitude1LSV3D2) = 1 then
        aLongitude1LSV3D2 := '0' + aLongitude1LSV3D2;
      if System.Pos('W', UpperCase(aLongitude1LSV3D)) > 0 then
        aLONSignLSV3D := '-'
      else
        aLONSignLSV3D := '';
      Result := aLONSignLSV3D + aLongitude1LSV3D1 + GlobLocaleFormats1.DecimalSeparator + aLongitude1LSV3D2;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '0';
      GlobSystemData1.aLogError1(E1, 'GetLongitudeString1');
      Exit;
    end;
  end;
end;

function GetEncrypted_SQL_Statement1(aValueLSV56: string): string;
var
  stringTempSQL1: string;
begin
  try
    Result := '';
    try
      stringTempSQL1 := 'SELECT CONVERT(NVARCHAR(MAX),DECRYPTBYPASSPHRASE(' +
                        QuotedStr(IntToStr(constCypher1)) +
                        ',[Name1])) AS [AVALUE1] FROM [' +
                        constSQL_Statements_Table1 +
                        '] WHERE ([Number1] = N' +
                        QuotedStr(aValueLSV56) +
                        ')' + stringGlobInstallationAndNewAndArx1;
      Result := GetStringValueFromQuery1(stringTempSQL1, 'AVALUE1');
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetEncrypted_SQL_Statement1');
      Exit;
    end;
  end;
end;

function GetStringValueFromAnyValue1(aField_Name1LV41D, aField_Name2LV41D, aTable_NameLV41D, aValueLV41D: string): string;
var
  stringTempSQL2: string;
begin
  try
    Result := '';
    try
      stringTempSQL2 := 'SELECT [' +
                        aField_Name1LV41D +
                        '] AS [AVALUE1] FROM [' +
                        aTable_NameLV41D +
                        '] WHERE ([' +
                        aField_Name2LV41D +
                        '] = N' +
                        QuotedStr(aValueLV41D) +
                        ')' + stringGlobInstallationAndNewAndArx1;
      Result := GetStringValueFromQuery1(stringTempSQL2, 'AVALUE1');
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetStringValueFromAnyValue1');
      Exit;
    end;
  end;
end;

function GetStringValueFromFloat1(aTempQueryLV4D: TMSQuery; stringFname4D: string; intPrecision1: integer): string;
begin
  try
    Result := '0';
    try
      if Trim(aTempQueryLV4D.FieldByName(stringFname4D).AsString) <> '' then
        Result := FloatToStrF(aTempQueryLV4D.FieldByName(stringFname4D).AsFloat, ffFixed, 18, intPrecision1);
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '0';
      GlobSystemData1.aLogError1(E1, 'GetStringValueFromFloat1');
      Exit;
    end;
  end;
end;

function GetNoOfMessagesWithThisID1(aMessage_ID1LSV4: string): integer;
var
  stringTempSQL45F: string;
begin
  try
    Result := 0;
    try
      stringTempSQL45F:='SELECT Count([Email_ID]) AS [NOOFEMAILS] FROM [' +
                        constTable_Name_21_22_23_Emails1 +
                        '] WHERE ([Number1] = N' +
                        QuotedStr(aMessage_ID1LSV4)+')' +
                        stringGlobInstallationAndNewAndArx1;
      Result := GetIntegerValueFromQuery1(stringTempSQL45F, 'NOOFEMAILS');
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 0;
      GlobSystemData1.aLogError1(E1, 'GetNoOfMessagesWithThisID1');
      Exit;
    end;
  end;
end;

function GetNoOfNewNotifications1: integer;
var
  stringTempSQL3: string;
begin
  try
    Result := 0;
    try
      if GlobLogon_User1.stringUser_ID1 <> '1' then
      begin
        stringTempSQL3 := 'SELECT Count(DISTINCT [User_To_Message_ID]) AS NOOFU2M1 FROM [' +
                          constTable_Name_119_User_To_Notifications1 +
                          '] WHERE ([Person_ID] = ' +
                          GlobLogon_User1.stringUser_ID1 +
                          ') AND ([' +
                          constStatusFieldTitle1 + '] = N' +
                          QuotedStr(constActiveStatusTitle1) +
                          ') AND ([Table_ID] = ' +
                          const122_ConversationRN1 + ')' +
                          stringGlobInstallationAndNewAndArx1;
        Result := GetIntegerValueFromQuery1(stringTempSQL3, 'NOOFU2M1');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 0;
      GlobSystemData1.aLogError1(E1, 'GetNoOfNewNotifications1');
      Exit;
    end;
  end;
end;

function GetNoOfPosRepsBetweenDates1(dtStart25, dtFinish25: TDateTime; stringMap_Ship_ID_List25: string): integer;
var
  stringTempSQL45F: string;
begin
  try
    Result := 0;
    try
      if (stringMap_Ship_ID_List25 <> '') and (dtStart25 <> NullDate) and (dtFinish25 <> NullDate) then
      begin
        stringTempSQL45F := 'SELECT Count([' +
                            constPosRepIndex_Name1 +
                            ']) AS [NOOFPOSREPS] FROM [' +
                            constPosRepTable_Name1 +
                            '] WHERE ([Unit_ID] IN (' +
                            stringMap_Ship_ID_List25 +
                            ')) AND (DATEADD(dd, 0, DATEDIFF(dd, 0, [Date1])) >= ' +
                            QuotedStr(FormatDateTime(constDBImportDateFormat1, dtStart25)) + ') ' +
                            'AND (DATEADD(dd, 0, DATEDIFF(dd, 0, [Date1])) <= ' +
                            QuotedStr(FormatDateTime(constDBImportDateFormat1, dtFinish25)) + ') ' +
                            stringGlobInstallationAndNewAndArx1;
        Result := GetIntegerValueFromQuery1(stringTempSQL45F, 'NOOFPOSREPS');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 0;
      GlobSystemData1.aLogError1(E1, 'GetNoOfPosRepsBetweenDates1');
      Exit;
    end;
  end;
end;

function GetNoOfGPSBetweenDates1(dtStart24, dtFinish24: TDateTime; stringMap_Ship_ID_List24: string): integer;
var
  stringTempSQL45F: string;
begin
  try
    Result := 0;
    try
      if (stringMap_Ship_ID_List24 <> '') and (dtStart24 <> NullDate) and (dtFinish24 <> NullDate) then
      begin
        stringTempSQL45F := 'SELECT Count([' +
                            constTable_Name_114_GPS_Position_ID +
                            ']) AS [NOOFPOSREPS] FROM [' +
                            constTable_Name_114_GPS_Positions1 +
                            '] WHERE ([Unit_ID] IN (' +
                            stringMap_Ship_ID_List24 +
                            ')) AND (DATEADD(dd, 0, DATEDIFF(dd, 0, [Date1])) >= ' +
                            QuotedStr(FormatDateTime(constDBImportDateFormat1, dtStart24)) + ') ' +
                            'AND (DATEADD(dd, 0, DATEDIFF(dd, 0, [Date1])) <= ' +
                            QuotedStr(FormatDateTime(constDBImportDateFormat1, dtFinish24)) + ') ' +
                            stringGlobInstallationAndNewAndArx1;
        Result := GetIntegerValueFromQuery1(stringTempSQL45F, 'NOOFPOSREPS');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 0;
      GlobSystemData1.aLogError1(E1, 'GetNoOfGPSBetweenDates1');
      Exit;
    end;
  end;
end;

function GetNoOfUserActiveTasks1: integer;
var
  stringTempSQL4: string;
begin
  try
    Result := 0;
    try
      if GlobLogon_User1.stringUser_ID1 <> '1' then
      begin
        stringTempSQL4:= 'SELECT Count(DISTINCT [' +
                         constTable_Name_45_Task_ID +
                         ']) AS NOOFU2M1 FROM [' +
                         constTable_Name_45_Tasks1 +
                         '] WHERE ([Person_ID] = ' +
                         GlobLogon_User1.stringUser_ID1 +
                         ') AND ([' + constStatusFieldTitle1 +
                         '] = N' +
                         QuotedStr(constActiveStatusTitle1) +
                         ')' +
                         stringGlobInstallationAndNewAndArx1;
        Result := GetIntegerValueFromQuery1(stringTempSQL4, 'NOOFU2M1');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 0;
      GlobSystemData1.aLogError1(E1, 'GetNoOfUserActiveTasks1');
      Exit;
    end;
  end;
end;

function GetTimeZoneDoubleFromString1(aTime_Zone_StringLSV34: string): Double;
var
  stringTemp1: string;
begin
  try
    Result := 1;
    stringTemp1 := aTime_Zone_StringLSV34;
    if Trim(stringTemp1) <> '' then
    begin
      stringTemp1 := Trim(stringTemp1);
      stringTemp1 := stringReplace(stringTemp1, ' ', '', [rfIgnoreCase, rfReplaceAll]);
      stringTemp1 := stringReplace(stringTemp1, '+', '', [rfIgnoreCase, rfReplaceAll]);
      if (System.Pos('.', stringTemp1) <> 0) or (System.Pos(',', stringTemp1) <> 0) then
      begin
        if (System.Pos('.', stringTemp1) <> 0) then stringTemp1 := Trim(System.Copy(stringTemp1, 1, System.Pos('.', stringTemp1) - 1));
        if (System.Pos(',', stringTemp1) <> 0) then stringTemp1 := Trim(System.Copy(stringTemp1, 1, System.Pos(',', stringTemp1) - 1));
      end;
      try
        Result := CustomStrToFloat1(stringTemp1);
      except
        Exit;
      end;
    end;
  except
    on E1: Exception do
    begin
      Result := 1;
      GlobSystemData1.aLogError1(E1, 'GetTimeZoneDoubleFromString1');
      Exit;
    end;
  end;
end;

function GetDetail_IDFromANamedField1(aTable_IDLSV3H, aText_Field_ValueSV3H, aField_NameLSV3H: string): string;
var
  stringTempSQL6, stringTable_Name1, stringIndex_Title1: string;
begin
  try
    Result := '';
    try
      if (aText_Field_ValueSV3H = 'N/A') then
      begin
        Result := '1';
      end
      else if (aText_Field_ValueSV3H <> '') then
      begin
        stringTable_Name1 := '';
        stringIndex_Title1 := '';
        stringTable_Name1 := GetRecord_Type_Detail1('Name1', aTable_IDLSV3H);
        stringIndex_Title1 := GetRecord_Type_Detail1('Number1', aTable_IDLSV3H);
        if (stringTable_Name1 <> '') and (stringIndex_Title1 <> '') then
        begin
          stringTempSQL6 := 'SELECT [' +
                            stringIndex_Title1 +
                            '] As [Record_IDX] FROM [' +
                            stringTable_Name1 +
                            '] WHERE (UPPER(LTRIM(RTRIM([' +
                            aField_NameLSV3H +
                            ']))) = N' +
                            QuotedStr(UpperCase(aText_Field_ValueSV3H)) +
                            ')' +
                            stringGlobInstallationAndNewAndArx1;
          Result := GetStringValueFromQuery1(stringTempSQL6, 'Record_IDX');
        end;
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetDetail_IDFromANamedField1');
      Exit;
    end;
  end;
end;

function GetDetail_IDFromAnIndexTitle1(aTable_IDLSV3H, aText_Field_ValueSV3H, aField_NameLSV3H: string): string;
var
  stringTempSQL7, stringTable_Name1, stringIndex_Title1: string;
begin
  try
    Result := '';
    try
      stringTable_Name1 := '';
      stringIndex_Title1 := '';
      stringTable_Name1 := GetRecord_Type_Detail1('Name1', aTable_IDLSV3H);
      stringIndex_Title1 := GetRecord_Type_Detail1('Number1', aTable_IDLSV3H);
      if (stringTable_Name1 <> '') then
      begin
        stringTempSQL7 := 'SELECT [' +
                          aField_NameLSV3H +
                          '] As [Record_IDX] FROM [' +
                          stringTable_Name1 +
                          '] WHERE (UPPER(LTRIM(RTRIM([' +
                          stringIndex_Title1 +
                          ']))) = N' +
                          QuotedStr(UpperCase(aText_Field_ValueSV3H)) +
                          ')' +
                          stringGlobInstallationAndNewAndArx1;
        Result := GetStringValueFromQuery1(stringTempSQL7, 'Record_IDX');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetDetail_IDFromAnIndexTitle1');
      Exit;
    end;
  end;
end;

function GetPerson_IDFromLogin_ID(aLogin_IDLVS66: string): string;
var
  stringTempSQL8: string;
begin
  try
    Result := '';
    try
      if (aLogin_IDLVS66 <> '') then
      begin
        stringTempSQL8 := 'SELECT [' +
                          constTable_Name_32_33_50_Personnel_ID +
                          '] FROM [' +
                          constTable_Name_32_33_50_Personnel1 +
                          '] WHERE (LOWER([Name1]) = N' +
                          QuotedStr(LowerCase(Trim(aLogin_IDLVS66))) +
                          ')' +
                          stringGlobInstallationAndNewAndArx1;
        Result := GetStringValueFromQuery1(stringTempSQL8, constTable_Name_32_33_50_Personnel_ID);
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetPerson_IDFromLogin_ID');
      Exit;
    end;
  end;
end;

function GetPerson_IDFromNameSurname1(stringPerson_Name1: string): string;
var
  stringTempSQL8, stringTempSQL9, stringFirstName1, stringLastName1, stringTemp8, stringTemp_Resu1: string;
begin
  try
    Result := '';
    try
      stringFirstName1 := '';
      stringLastName1 := '';
      stringTemp_Resu1 := '';
      if True then
        if System.Pos(' ', stringPerson_Name1) <> 0 then
        begin
          stringFirstName1 := Trim(System.Copy(stringPerson_Name1, 1, System.Pos(' ', stringPerson_Name1)));
          stringTemp8 := Trim(System.Copy(stringPerson_Name1, System.Pos(' ', stringPerson_Name1) + 1, Length(stringPerson_Name1)));
          if System.Pos(',', stringTemp8) <> 0 then
          begin
            stringLastName1 := Trim(System.Copy(stringTemp8, 1, System.Pos(',', stringTemp8) - 1));
          end
          else
          begin
            if System.Pos(' ', stringTemp8) <> 0 then
            begin
              stringLastName1 := Trim(System.Copy(stringTemp8, 1, System.Pos(' ', stringTemp8)));
            end
            else
            begin
              stringLastName1 := Trim(System.Copy(stringTemp8, 1, Length(stringTemp8)));
            end;
          end;
        end;
      if (stringFirstName1 <> '') AND (stringLastName1 <> '') then
      begin
        stringTempSQL8 := 'SELECT [' +
                          constTable_Name_32_33_50_Personnel_ID +
                          '] FROM [' +
                          constTable_Name_32_33_50_Personnel1 +
                          '] WHERE (UPPER(LTRIM(RTRIM([First_Name]))) = N' +
                          QuotedStr(UpperCase(stringFirstName1)) +
                          ') AND (UPPER(LTRIM(RTRIM([Last_Name]))) = N' +
                          QuotedStr(UpperCase(stringLastName1)) +
                          ')' +
                          stringGlobInstallationAndNewAndArx1;
        stringTemp_Resu1 := GetStringValueFromQuery1(stringTempSQL8, constTable_Name_32_33_50_Personnel_ID);
        if stringTemp_Resu1 <> '' then
          Result := stringTemp_Resu1
        else
        begin
         stringTempSQL9 := 'SELECT [' +
                           constTable_Name_32_33_50_Personnel_ID +
                           '] FROM [' +
                           constTable_Name_32_33_50_Personnel1 +
                           '] WHERE (UPPER(LTRIM(RTRIM([First_Name]))) = N' +
                           QuotedStr(UpperCase(stringLastName1)) +
                           ') AND (UPPER(LTRIM(RTRIM([Last_Name]))) = N' +
                           QuotedStr(UpperCase(stringFirstName1)) +
                           ')' + stringGlobInstallationAndNewAndArx1;
         Result := GetStringValueFromQuery1(stringTempSQL9, constTable_Name_32_33_50_Personnel_ID);
        end;
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetPerson_IDFromNameSurname2');
      Exit;
    end;
  end;
end;

function GetPersonNameSurnameFromPerson_ID1(aPerson_IDLVS66: string): string;
var
  stringTempSQL10: string;
begin
  try
    Result := '';
    try
      stringTempSQL10 := 'SELECT LTRIM(ISNULL([First_Name],'''')+'' ''+[Last_Name]) AS [NAMESURNAME1] FROM [' +
                         constTable_Name_32_33_50_Personnel1 +
                         '] WHERE ([' +
                         constTable_Name_32_33_50_Personnel_ID +
                         '] = ' +
                         aPerson_IDLVS66 +
                         ')' +
                         stringGlobInstallationAndNewAndArx1;
      Result := GetStringValueFromQuery1(stringTempSQL10, 'NAMESURNAME1');
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetPersonNameSurnameFromPerson_ID1');
      Exit;
    end;
  end;
end;

function GetPersonLoginDateTimeFromPerson_ID(Person_IDLVS66: string): string;
var
  stringTempSQL11: string;
begin
  try
    Result := '';
    try
      if (Person_IDLVS66 <> '') then
      begin
        stringTempSQL11 := 'SELECT  MAX([Date1]) AS LOGINDATETIME1 FROM [' +
                           constTable_Name_71_User_Logins1 +
                           '] WHERE ([Person_ID] = ' +
                           Person_IDLVS66 +
                           ')' +
                           stringGlobInstallationAndNewAndArx1;
        Result := GetDateTimeStrValueFromQuery1(stringTempSQL11, 'LOGINDATETIME1', 'dd/mm/yyyy hh:mm');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetPersonLoginDateTimeFromPerson_ID');
      Exit;
    end;
  end;
end;

function GetRecord_Type_Detail1(aDetail_Field_NameIDLSV843, aRecord_Type_IDLSV843: string): string;
var
  stringTempSQL12: string;
begin
  try
    Result := '';
    try
      if FormDM1.Database_Connection_SDAC.Connected = True then
      begin
         if ((aDetail_Field_NameIDLSV843 <> 'Remarks1')
            AND (aDetail_Field_NameIDLSV843 <> 'Name1')
              AND (aDetail_Field_NameIDLSV843 <> 'Number1')
                AND (aDetail_Field_NameIDLSV843 <> 'Number2')
                  AND (aDetail_Field_NameIDLSV843 <> 'Number3')
                    AND (aDetail_Field_NameIDLSV843 <> 'Number4')
                      AND (aDetail_Field_NameIDLSV843 <> 'Number5')) then
            begin
               stringTempSQL12 := 'SELECT [' +
                                  aDetail_Field_NameIDLSV843 +
                                  '] FROM [' +
                                  constTable_Name_38_Record_Types1 +
                                  '] WHERE ([Record_Type_ID] = ' +
                                  aRecord_Type_IDLSV843+ ')' +
                                  stringGlobInstallationAndNewAndArx1;
            end else
            begin
              stringTempSQL12 :=  'SELECT CONVERT(NVARCHAR(MAX),DECRYPTBYPASSPHRASE(' +
                                  QuotedStr(IntToStr(constCypher1)) +
                                  ',[' +
                                  aDetail_Field_NameIDLSV843 +
                                  '])) AS [' +
                                  aDetail_Field_NameIDLSV843 +
                                  '] FROM [' +
                                  constTable_Name_38_Record_Types1 +
                                  '] WHERE ([Record_Type_ID] = ' +
                                  aRecord_Type_IDLSV843+ ')' +
                                  stringGlobInstallationAndNewAndArx1;
            end;
        Result := GetStringValueFromQuery1(stringTempSQL12, aDetail_Field_NameIDLSV843);
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetRecord_Type_Detail1');
      Exit;
    end;
  end;
end;

function GetRecord_Type_Detail2(aDetail_Field_NameIDLSV843, aRecord_Type_Index_IDLSV843: string): string;
var
  stringTempSQL14: string;
begin
  try
    Result := '';
    try
      if FormDM1.Database_Connection_SDAC.Connected = True then
      begin
        if (aDetail_Field_NameIDLSV843 <> 'Remarks1')
           and (aDetail_Field_NameIDLSV843 <> 'Name1')
             and (aDetail_Field_NameIDLSV843 <> 'Number1')
               and (aDetail_Field_NameIDLSV843 <> 'Number2')
                and (aDetail_Field_NameIDLSV843 <> 'Number3')
                  and (aDetail_Field_NameIDLSV843 <> 'Number4')
                    and (aDetail_Field_NameIDLSV843 <> 'Number5') then
        begin
          stringTempSQL14 :=  'SELECT [' +
                              aDetail_Field_NameIDLSV843 +
                              '] FROM [' +
                              constTable_Name_38_Record_Types1 +
                              '] WHERE (CONVERT(NVARCHAR(MAX),DECRYPTBYPASSPHRASE(' +
                              QuotedStr(IntToStr(constCypher1)) +
                              ',[Number1])) = N' +
                              QuotedStr(aRecord_Type_Index_IDLSV843) +
                              ')' +
                              stringGlobInstallationAndNewAndArx1;
        end
        else
        begin
           stringTempSQL14 :=  'SELECT CONVERT(NVARCHAR(MAX),DECRYPTBYPASSPHRASE(' +
                               QuotedStr(IntToStr(constCypher1)) +
                               ',[' +
                               aDetail_Field_NameIDLSV843 +
                               '])) AS [' +
                               aDetail_Field_NameIDLSV843 +
                               '] FROM [' +
                               constTable_Name_38_Record_Types1 +
                               '] WHERE (CONVERT(NVARCHAR(MAX),DECRYPTBYPASSPHRASE(' +
                               QuotedStr(IntToStr(constCypher1)) +
                               ',[Number1])) = N' +
                               QuotedStr(aRecord_Type_Index_IDLSV843) +
                               ')' +
                               stringGlobInstallationAndNewAndArx1;
        end;
        Result := GetStringValueFromQuery1(stringTempSQL14, aDetail_Field_NameIDLSV843);
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetRecord_Type_Detail2');
      Exit;
    end;
  end;
end;

function GetColourNumberFromColour_ID1(stringColour_ID1: string): integer;
var
  stringTempSQL15: string;
begin
  try
    Result := 14120960;
    try
      if FormDM1.Database_Connection_SDAC.Connected = True then
      begin
        stringTempSQL15 := 'SELECT [Colour1] FROM [Colours1] ' +
                           'WHERE ([Colour_ID] = ' + stringColour_ID1 + ')' +
                           stringGlobInstallationAndNewAndArx1;
        Result := GetIntegerValueFromQuery1(stringTempSQL15, 'Colour1');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 14120960;
      GlobSystemData1.aLogError1(E1, 'GetColourNumberFromColour_ID1');
      Exit;
    end;
  end;
end;

function GetRecord_Type_Colour1(stringRecord_Type_ID1: string): integer;
var
  stringTempSQL15, intColour_ID1: string;
begin
  try
    Result := 14120960;
    try
      if FormDM1.Database_Connection_SDAC.Connected = True then
      begin
        stringTempSQL15 := 'SELECT [Colour_ID] FROM [' +
                           constTable_Name_38_Record_Types1 +
                           '] WHERE ([Record_Type_ID] = ' +
                           stringRecord_Type_ID1 + ')' +
                           stringGlobInstallationAndNewAndArx1;
        intColour_ID1 := GetStringValueFromQuery1(stringTempSQL15, 'Colour_ID');
        if intColour_ID1 = '' then
          intColour_ID1 := '1';
        Result := GetColourNumberFromColour_ID1(intColour_ID1);
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 14120960;
      GlobSystemData1.aLogError1(E1, 'GetRecord_Type_Colour1');
      Exit;
    end;
  end;
end;

function GetMaxNewlyCreatedID1(aCurrentDataGridLV317J: TaCurrentDataTableStructure): string;
var
  stringTempSQL15: string;
begin
  try
    Result := '';
    try
      stringTempSQL15 :=  'SELECT Max([' +
                          aCurrentDataGridLV317J.stringRecord_Index_Title1 +
                          ']) AS [AIDX] from [' +
                          aCurrentDataGridLV317J.stringTable_Title1 +
                          '] WHERE ([Installation_ID] = ' +
                          GlobSystemData1.stringInstallation_ID +
                          ') AND ([Arx] = 0)';
      Result := GetStringValueFromQuery1(stringTempSQL15, 'AIDX');
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetMaxNewlyCreatedID1');
      Exit;
    end;
  end;
end;

function GetTopPanelPosition1L1(aPanel: TcxGroupBox): integer;
begin
  Result := aPanel.Top + aPanel.Height;
end;

function GetFileInfoSize1(aStrPath1: string): string;
var
  aSearchRec1: TSearchRec;
  intSize1: integer;

  function aGetSizeDescription1(const IntSize: Int64): string;
  begin
    try
      if IntSize < 1024 then
        Result := IntToStr(IntSize) + ' bytes'
      else
      begin
        if IntSize < (1024 * 1024) then
          Result := FormatFloat('####0.##', IntSize / 1024) + ' Kb'
        else if IntSize < (1024 * 1024 * 1024) then
          Result := FormatFloat('####0.##', IntSize / 1024 / 1024) + ' Mb'
        else
          Result := FormatFloat('####0.##', IntSize / 1024 / 1024 / 1024) + ' Gb';
      end;
    except
      on E1: Exception do
      begin
        Result := '';
        GlobSystemData1.aLogError1(E1, 'aGetFileInfoSize1.1');
        Exit;
      end;
    end;
  end;

begin
  try
    if Trim(aStrPath1) = '' then
    begin
      Result := '';
    end
    else
    begin
      try
        FindFirst(aStrPath1, 0, aSearchRec1);
        intSize1 := aSearchRec1.Size;
        Result := aGetSizeDescription1(intSize1);
      finally
        FindClose(aSearchRec1);
      end;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'aGetFileInfoSize1');
      Exit;
    end;
  end;
end;

function GetFileInfoCreateDate1(aStrPath1: string): string;
var
  aSearchRec1: TSearchRec;
  aDateTime1: TDateTime;
begin
  try
    if Trim(aStrPath1) = '' then
    begin
      Result := '';
    end
    else
    begin
      try
        FindFirst(aStrPath1, 0, aSearchRec1);
        try
          aDateTime1 := FileDateToDateTime(aSearchRec1.Time);
        except
          aDateTime1 := Now();
        end;
      finally
        FindClose(aSearchRec1);
      end;
      Result := FormatDateTime(constDBImportDateTimeSecFormat1, aDateTime1);
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'aGetFileInfoSize1');
      Exit;
    end;
  end;
end;

function GetShip_ColourFromShip_ID1(stringShip_ID1: string): integer;
var
  stringTempSQL16: string;
begin
  try
    Result := 1;
    try
      if (stringShip_ID1 <> '') then
      begin
        stringTempSQL16 :=  'SELECT [Colour_ID] FROM [' +
                            constTable_Name_46_Units1 +
                            '] WHERE ([' +
                            constTable_Name_46_Unit_ID +
                            '] = ' +
                            QuotedStr(stringShip_ID1) +
                            ')' +
                            stringGlobInstallationAndNewAndArx1;
        Result := GetIntegerValueFromQuery1(stringTempSQL16, 'Colour_ID');
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 1;
      GlobSystemData1.aLogError1(E1, 'GetShip_ColourFromShip_ID1');
      Exit;
    end;
  end;
end;

function GetHeightOfTaskBar1: integer;
var
  handleTaskBar1: HWND;
  rectR1: TRect;
begin
  try
    Result := 10;
    try
      handleTaskBar1 := FindWindow('Shell_TrayWnd', Nil);
      if handleTaskBar1 <> 0 then
        GetWindowRect(handleTaskBar1, rectR1);
      Result := rectR1.bottom - rectR1.top;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 10;
      GlobSystemData1.aLogError1(E1, 'GetHeightOfTaskBar1');
      Exit;
    end;
  end;
end;

function GetMaxVoyDate1(aDate_IDLSV23TR4, aUnit_IDLSV23TR4: string): TDateTime;
var
  stringTempSQL17: string;
begin
  try
    Result := NullDate;
    try
      stringTempSQL17 := 'SELECT [' +
                         aDate_IDLSV23TR4 +
                         '] FROM [' +
                         constTable_Name_49_Voyages1 +
                         '] WHERE ([' +
                         constTable_Name_49_Voyage_ID +
                         '] = (SELECT MAX([' +
                         constTable_Name_49_Voyage_ID +
                         ']) AS [Max_Voyage_ID] FROM [' +
                         constTable_Name_49_Voyages1 +
                         '] AS [V1] WHERE ([Unit_ID] = ' +
                         aUnit_IDLSV23TR4 +
                         ')' + stringGlobInstallationAndNewAndArx1 + ')';
      Result := GetDateTimeValueFromQuery1(stringTempSQL17, aDate_IDLSV23TR4);
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := NullDate;
      GlobSystemData1.aLogError1(E1, 'aGetMaxVoyDate1');
      Exit;
    end;
  end;
end;

function GetMaxVoyDoubleDataFromUnitID1(aField_NameLSV23TR4, aUnit_IDLSV23TR4: string): Double;
var
  stringTempSQL18: string;
begin
  try
    Result := 0;
    try
      stringTempSQL18 := 'SELECT [' +
                         aField_NameLSV23TR4 +
                         '] FROM [' +
                         constTable_Name_49_Voyages1 +
                         '] WHERE [' +
                         constTable_Name_49_Voyage_ID +
                         '] = (SELECT MAX([' +
                         constTable_Name_49_Voyage_ID +
                         ']) AS [Max_Voyage_ID]  FROM [' +
                         constTable_Name_49_Voyages1 +
                         '] AS [V1] WHERE ([Unit_ID] = ' +
                         aUnit_IDLSV23TR4 +
                         ')' + stringGlobInstallationAndNewAndArx1 + ')';
      Result := GetFloatValueFromQuery1(stringTempSQL18, aField_NameLSV23TR4);
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := 0;
      GlobSystemData1.aLogError1(E1, 'aGetMaxVoyDataFromUnitID1');
      Exit;
    end;
  end;
end;

function GetNoOfConnectedRecords1(aDaughter_IDIDLSV87F, aParent_IDLSV87F, aParent_IndexLSV87F: string): integer;
var
  stringTempSQL19: string;
  aCurrentDetailsLV1, aCurrentDetailsLV2: TaCurrentDataDetails;
begin
  try
    aCurrentDetailsLV1 := TaCurrentDataDetails.Create;
    aCurrentDetailsLV1.InitCurrentDataDetails;
    aCurrentDetailsLV2 := TaCurrentDataDetails.Create;
    aCurrentDetailsLV2.InitCurrentDataDetails;
    try
      Result := 0;
      try
        aCurrentDetailsLV1.aRecord_Type_ID1 := aDaughter_IDIDLSV87F;
        aCurrentDetailsLV2.aRecord_Type_ID1 := aParent_IDLSV87F;
        if aCurrentDetailsLV1.PopulateCurrent_RecordType1(aCurrentDetailsLV1.aRecord_Type_ID1) = True then
        begin
          if aCurrentDetailsLV2.PopulateCurrent_RecordType1(aCurrentDetailsLV2.aRecord_Type_ID1) = True then
          begin
            stringTempSQL19 :=  'SELECT Count([' +
                                aCurrentDetailsLV1.stringRecord_Index_Title1 +
                                ']) AS NOFDRS1 FROM [' +
                                aCurrentDetailsLV1.stringTable_Title1 +
                                '] WHERE ([Installation_ID] = ' +
                                GlobSystemData1.stringInstallation_ID +
                                ') AND ([Arx] = 0) AND ([New] = 0) ';
            if (aCurrentDetailsLV1.aRecord_Type_ID1 = const047_UpdatesLogRN1)
                or (aCurrentDetailsLV1.aRecord_Type_ID1 = const073_User_ReadRN1)
                or (aCurrentDetailsLV1.aRecord_Type_ID1 = const021_EmailRN1)
                or (aCurrentDetailsLV1.aRecord_Type_ID1 = const122_ConversationRN1)
                or (aCurrentDetailsLV1.aRecord_Type_ID1 = const007_AttachmentRN1)
                or (aCurrentDetailsLV1.aRecord_Type_ID1 = const012_CommentRN1) then
            begin
              stringTempSQL19 := stringTempSQL19 +  'AND ([Table_ID] = ' +
                                                    aCurrentDetailsLV2.aRecord_Type_ID1 +
                                                    ') AND ([Table_Index_ID] = ' +
                                                    aParent_IndexLSV87F + ')';
            end
            else if (aCurrentDetailsLV2.aRecord_Type_ID1 = const061_DayRepRN1) then
            begin
              stringTempSQL19 := stringTempSQL19 +  'AND ([' + constSwitchFieldTitle1 +
                                                    '] = ' +
                                                    aCurrentDetailsLV1.aSwitch1 +
                                                    ') AND ([' +
                                                    constPosRepIndex_Name2 +
                                                    '] = ' +
                                                    aParent_IndexLSV87F + ')';
            end
            else
            begin
              if (aCurrentDetailsLV2.aRecord_Type_ID1 = const017_DepartmentRN1)
                 or (aCurrentDetailsLV2.aRecord_Type_ID1 = const059_DepartmentRN2) then
              begin
                stringTempSQL19 := stringTempSQL19 + 'AND ([' + constStatusFieldTitle1 +'] = N' + QuotedStr(constActiveStatusTitle1) + ')';
              end;
              stringTempSQL19 := stringTempSQL19 + 'AND ([' + constSwitchFieldTitle1 +'] = ' +
              aCurrentDetailsLV1.aSwitch1 + ') ' +
              'AND ([' + aCurrentDetailsLV2.stringRecord_Index_Title1 +
              '] = ' + aParent_IndexLSV87F + ')';
            end;
            if (aDaughter_IDIDLSV87F = const014_ContactRN1)
              or (aDaughter_IDIDLSV87F = const079_Follow_UpRN1)
               or (aDaughter_IDIDLSV87F = const078_PerformanceRN1)
                or (aDaughter_IDIDLSV87F = const120_UtilisationRN1)
                 or (aDaughter_IDIDLSV87F = const076_Miscellaneous_DataRN1) then
              Result := 1
            else
              Result := GetIntegerValueFromQuery1(stringTempSQL19, 'NOFDRS1');
          end;
        end;
      except
        Exit;
      end;
    except
      on E1: Exception do
      begin
        Result := 0;
        GlobSystemData1.aLogError1(E1, 'GetNoOfConnectedRecords1');
        Exit;
      end;
    end;
  finally
   if Assigned(aCurrentDetailsLV1) = True then FreeAndNil(aCurrentDetailsLV1);
   if Assigned(aCurrentDetailsLV2) = True then FreeAndNil(aCurrentDetailsLV2);
  end;
end;

function GetUpdateStringWithIntegerFromCurrentDataDetailFromIndex1(aCurrentDetailsLV88LPZ: TaRecordTypeStrings; aFied_NameLSV88LPZ, aField_ValueLSV88LPZ: string): string;
begin
  try
    Result := '';
    try
      Result := 'UPDATE [' +
                aCurrentDetailsLV88LPZ.stringTable_Title1 +
                '] Set [' +
                aFied_NameLSV88LPZ +
                '] = ' +
                aField_ValueLSV88LPZ +
                ' WHERE (' +
                aCurrentDetailsLV88LPZ.stringRecord_Index_Title1 +
                ' = ' +
                aCurrentDetailsLV88LPZ.aRecord_Detail_Index1 + ')';
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetUpdateStringWithIntegerFromCurrentDataDetailFromIndex1');
      Exit;
    end;
  end;
end;

function GetUpdateStringWithStringFromCurrentDataDetailFromIndex1(aCurrentDetailsLV88LPZ: TaRecordTypeStrings; aFied_NameLSV88LPZ, aField_ValueLSV88LPZ: string): string;
begin
  try
    Result := '';
    try
      Result := 'UPDATE [' +
                aCurrentDetailsLV88LPZ.stringTable_Title1 +
                '] Set [' +
                aFied_NameLSV88LPZ +
                '] = N' +
                QuotedStr(aField_ValueLSV88LPZ) +
                ' WHERE (' +
                aCurrentDetailsLV88LPZ.stringRecord_Index_Title1 +
                ' = ' +
                aCurrentDetailsLV88LPZ.aRecord_Detail_Index1 + ')';
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetUpdateStringWithStringFromCurrentDataDetailFromIndex1');
      Exit;
    end;
  end;
end;

function GetTilesButtonsWidth1(aTileControlLV34 : TdxTileControl): Integer;
var
  intLoopCounter3, intDelimiter1, intWidth1: Integer;
begin
  try
    Result := 0;
    try
      intWidth1 := 0;
      if aTileControlLV34.Items.Count > 0 then
      begin
        intDelimiter1 := aTileControlLV34.OptionsView.ItemIndent;
        for intLoopCounter3 := 0 to aTileControlLV34.Items.Count - 1 do
        begin
          intWidth1 := intWidth1 + aTileControlLV34.OptionsView.ItemWidth + intDelimiter1;
        end;
        Result := intWidth1 + 5 * GlobSystemData1.intDPIFactor1;
      end;
    except
      Exit;
    end;
    {SaveMsgToAFile1(constLogFolderName1, '', IntToStr(aWidth1) + '.' + IntToStr(FormMain.MainData1TilesButtons1.OptionsView.ItemWidth) + '.' +'.GetMainData1TilesButtonsWidth1.Error');}
  except
    on E1: Exception do
    begin
      Result := 0;
      GlobSystemData1.aLogError1(E1, 'GetTilesButtonsWidth1');
      Exit;
    end;
  end;
end;

function GetVoyage_IDFromShipData_Ship_IDAndVoyageNumber1(aShipDataLV34D: TaShipData): string;
var
  stringTempSQL20: string;
begin
  try
    Result := '';
    try
      if (aShipDataLV34D.stringShipID1 <> '') AND (aShipDataLV34D.stringVoyNo1 <> '') then
      begin
        stringTempSQL20 :=  'SELECT [' +
                            constTable_Name_49_Voyage_ID +
                            '] FROM [' +
                            constTable_Name_49_Voyages1 +
                            '] WHERE ([Unit_ID] = ' +
                            aShipDataLV34D.stringShipID1 +
                            ') AND ([Number1] = ' +
                            QuotedStr(aShipDataLV34D.stringVoyNo1) +
                            ')' +
                            stringGlobInstallationAndNewAndArx1;
        Result := GetStringValueFromQuery1(stringTempSQL20, constTable_Name_49_Voyages1);
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetVoyage_IDFromShipData_Ship_IDAndVoyageNumber1');
      Exit;
    end;
  end;
end;

function GetShip_IDFromShipData_IMO_Number1(aShipDataLV34D: TaShipData): string;
var
  stringTempSQL21: string;
begin
  try
    Result := '';
    try
      if (aShipDataLV34D.stringImoNo1 <> '') then
      begin
         stringTempSQL21 :=  'SELECT [' +
                             constTable_Name_46_Unit_ID +
                             '] FROM [' +
                             constTable_Name_46_Units1 +
                             '] WHERE ([Number1] = ' +
                             QuotedStr(aShipDataLV34D.stringImoNo1)+
                             ')' + stringGlobInstallationAndNewAndArx1;
        Result := GetStringValueFromQuery1(stringTempSQL21, constTable_Name_46_Unit_ID);
      end;
    except
      Exit;
    end;
  except
    on E1: Exception do
    begin
      Result := '';
      GlobSystemData1.aLogError1(E1, 'GetShip_IDFromShipData_IMO_Number1');
      Exit;
    end;
  end;
end;


end.
