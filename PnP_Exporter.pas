Program PnP_Exporter;
var
    PartsList   : TStringList;
    MountList   : TStringList;
    SizeList    : TStringList;
    MountVarList : TStringList;
    FeederList  : TStringList;
    SelectedLayer : TLayer;

{..............................................................................}
Function UnitToString(U : TUnit) : TPCBString;
Begin
    Result := '';
    Case U of
       eImperial : Result := 'Imperial (mil)';
       eMetric   : Result := 'Metric (mm)';
    End;
End;
{..............................................................................}

{..............................................................................}
Function BoolToString(B : Boolean) : TPCBString;
Begin
    Result := 'False';
    If B Then Result := 'True';
End;
{..............................................................................}

{..............................................................................}
Function RemoveIllegalChars(S : TPCBString) : TPCBString;
Var
    I : Integer;
    Asc : Integer;
Begin
    If Not VarIsNull(S) Then
    Begin
        For I := 1 to Length(S) Do
        Begin
            Asc := Ord(S[I]);
            If Asc < 48 Then
                S[I] := '_'
            else If (Asc > 57) and (Asc < 65) Then
                S[I] := '_'
            else If (Asc > 90) and (Asc < 95) Then
                S[I] := '_'
            else If Asc = 96 Then
                S[I] := '_'
            else If Asc > 122 Then
                S[I] := '_'
            else
                S[I] := Chr(Asc);
        End;
    End;
    Result := S;
End;
{..............................................................................}

{..............................................................................}
Function RemovePath(S : TPCBString) : TPCBString;
Var
    Count : Integer;
    I : Integer;
    
Begin
    Count := 0;
    If Not VarIsNull(S) Then
    Begin
        For I := 1 to Length(S) Do
            Begin
                If (Ord(S[I]) = 92) Then Count := I;
            End;
        Delete(S, 1, Count);
        Result := S;
    End;
End;
{..............................................................................}

{..............................................................................}
Procedure Query_OutlinePerimeter(Var BR : TCoordRect; Var OtherUnit : TUnit; Var Edge_Str : TPCBString; Var Area_Str : TPCBString);
Var
    PCB_Board : IPCB_Board;
    I,J       : Integer;
    Del_X     : Real;
    Del_Y     : Real;
    Del_Arc   : Real;
    Perimeter : Real;
    Arc_Angle : TAngle;

Begin
    PCB_Board       := PCBServer.GetCurrentPCBBoard;
    If PCB_Board     = Nil Then Exit;

    PCB_Board.BoardOutline.Invalidate;
    PCB_Board.BoardOutline.Rebuild;
    PCB_Board.BoardOutline.Validate;
    BR := PCB_Board.BoardOutline.BoundingRectangle;

    // Set OtherUnit to "opposite" of PCB_Board.DisplayUnit
    If PCB_Board.DisplayUnit = eImperial Then
        OtherUnit := eMetric
    Else
        OtherUnit := eImperial;

    // Determine length of perimeter.
    // Note that the first vertex of a polygon is the same as the last vertex.
    Perimeter := 0;
    For I := 0 To PCB_Board.BoardOutline.PointCount - 1 Do
    Begin
       If I = PCB_Board.BoardOutline.PointCount - 1 Then
           J := 0
       Else
           J := I + 1;
{
We want to calculate (PCB_Board.BoardOutline.Segments[I+1].vx - PCB_Board.BoardOutline.Segments[I].vx) for all
values of I from 0 to PCB_Board.BoardOutline.PointCount - 2; when I = PCB_Board.BoardOutline.PointCount - 1,
PCB_Board.BoardOutline.Segments[PCB_Board.BoardOutline.PointCount].vx is dodgy, hence the use of the J variable
in addition to the I variable.
}
        If PCB_Board.BoardOutline.Segments[I].Kind = ePolySegmentLine Then
        Begin
            Del_X := CoordToMils(PCB_Board.BoardOutline.Segments[J].vx
                               - PCB_Board.BoardOutline.Segments[I].vx);

            Del_Y := CoordToMils(PCB_Board.BoardOutline.Segments[J].vy
                               - PCB_Board.BoardOutline.Segments[I].vy);

            Perimeter := Perimeter + Sqrt((Del_X * Del_X) + (Del_Y * Del_Y));
        End
        Else
        Begin
            Arc_Angle := PCB_Board.BoardOutline.Segments[I].Angle2
                       - PCB_Board.BoardOutline.Segments[I].Angle1;

            If Arc_Angle < 0 Then Arc_Angle := Arc_Angle + 360;

            Del_Arc   := CoordToMils(PCB_Board.BoardOutline.Segments[I].Radius)
                       * Degrees2Radians(Arc_Angle);

            Perimeter := Perimeter + Del_Arc;
        End;
    End;

    // Construct perimeter-reporting string - this is reported in units of both inches and cm.
    // If the current value of the Measurement Unit is Imperial, the former is listed first,
    // then the latter within brackets; otherwise (i.e. the current value of the Measurement Unit
    // is Metric), the latter is listed first, then the former within brackets. (1 inch = 2.54 cm)
    If PCB_Board.DisplayUnit = eImperial Then
    Begin
        Edge_Str := FloatToStr(Perimeter * 0.001)   + ' inch' + '  ('
                  + FloatToStr(Perimeter * 0.00254) + ' cm)';
    End
    Else
    Begin
        Edge_Str := FloatToStr(Perimeter * 0.00254) + ' cm' + '  ('
                  + FloatToStr(Perimeter * 0.001)   + ' inch)';
    End;

    // Construct area-reporting string  - this is reported in units of both inch^2 and cm^2.
    // If the current value of the Measurement Unit is Imperial, the former is listed first,
    // then the latter within brackets; otherwise (i.e. the current value of the Measurement
    // Unit is Metric), the latter is listed first, then the former within brackets.
    // The AreaSize property is returned in units of 1 x 10E-14 inch^2, which is subsequently
    // scaled as required. (1 inch = 2.54 cm, so 1 inch^2 = (2.54)^2 cm^2 = 6.4516 cm^2.)
    If PCB_Board.DisplayUnit = eImperial Then
    Begin
        Area_Str := FloatToStr(PCB_Board.BoardOutline.AreaSize / (10000000 * 10000000))
                  + ' inch^2' + '  ('
                  + FloatToStr(PCB_Board.BoardOutline.AreaSize * 6.4516 / (10000000 * 10000000))
                  + ' cm^2)';
    End
    Else
    Begin
        Area_Str := FloatToStr(PCB_Board.BoardOutline.AreaSize * 6.4516 / (10000000 * 10000000))
                  + ' cm^2' + '  ('
                  + FloatToStr(PCB_Board.BoardOutline.AreaSize / (10000000 * 10000000))
                  + ' inch^2)';
    End;
End;
{..............................................................................}

{..............................................................................}
Procedure IterateComponents(BoardFile : TString);
Var
    Board         : IPCB_Board;
    Comp          : IPCB_Component;
    Iterator      : IPCB_BoardIterator;

    PartsIndex    : TStringList;


    I             : Integer;
    J             : Integer;
    K             : Integer;
    Added         : Integer;
    S             : TString;
    CompNum       : Integer;
    Count         : Integer;
    FeedPos       : Float;
    PlatePos      : Float;
    FeedOff       : Integer;
    IncCount      : Integer;

Begin
    // Open the board from the embedded array
    Board := PCBServer.GetPCBBoardByPath(BoardFile);
    If Board = Nil Then Exit;

    // Create the iterator that will look for components only
    Iterator := Board.BoardIterator_Create;
    Iterator.AddFilter_ObjectSet(MkSet(eComponentObject));
    Iterator.AddFilter_LayerSet(MkSet(SelectedLayer));
    Iterator.AddFilter_Method(eProcessAll);

    // Search for component objects and get their info
    PartsList  := TStringList.Create;
    MountList := TStringList.Create;
    PartsIndex := TStringList.Create;
    SizeList := TStringList.Create;
    MountVarList := TStringList.Create;
    FeederList := TStringList.Create;

    // Make the parts list
    Comp := Iterator.FirstPCBObject;
    Count := 1;
    IncCount := 0;
    While (Comp <> Nil) Do
    Begin
        S := '  <Component No="' + IntToStr(Count) + '" Name="' + RemoveIllegalChars(Comp.SourceLibReference) + '" Comment="' + RemoveIllegalChars(Comp.Pattern)
        + '" ID="' + IntToStr(Count - 1) + '" DatabaseNo="0" FixedComp="0">';
        // check if this component is already in the parts list, then add it if not
        Added := 0;

        For K := 0 to CheckListBox1.Items.Count - 1 Do
        Begin
             If (CompareStr(CheckListBox1.Items.Strings[K], Comp.SourceLibReference) = 0) Then IncCount := K;
        End;

        For K := 0 to PartsIndex.Count - 1 Do
        Begin
             If (CompareStr(PartsIndex[K], RemoveIllegalChars(Comp.SourceLibReference)) = 0) then Added := 1;
        End;

        If (Added = 0) And (CheckListBox1.Checked[IncCount]) Then
        Begin
            PartsIndex.Add(RemoveIllegalChars(Comp.SourceLibReference));
            SizeList.Add('          <Size  X="'+ FloatToStr(CoordToMMs(Comp.BoundingRectangleForSelection.right - Comp.BoundingRectangleForSelection.left))
              + '" Y="' + FloatToStr(CoordToMMs(Comp.BoundingRectangleForSelection.top - Comp.BoundingRectangleForSelection.bottom))
              + '" Z="'+ FloatToStr(CoordToMMs(Comp.Height)) +'"/>');
            PartsList.Add(S);
            If Count = 1 Then FeedPos := -66.175;
            If Count < 38 Then
            Begin
                FeedPos := FeedPos + 16.0;
                PlatePos := -34.866;
                FeedOff := 0;
            End;
            If Count = 38 Then
            Begin
                FeedPos := 318.232;
                PlatePos := 694.4;
            End;
            If Count > 37 Then
            Begin
                FeedPos := FeedPos - 16.0;
                PlatePos := 694.4;
                FeedOff := 89;
            End;
            FeederList.Add('            <Feeder Machine="6535" Number="' + IntToStr(Count + FeedOff) + '" Definition="0" X="' + FloatToStr(FeedPos) + '" Y="' + FloatToStr(FeedPos) + '" Opt="0" Group="0" ID="' + IntToStr(Count - 1) + '" CompRef="' + IntToStr(Count - 1) + '" AltFeederRef="-0000001"/>');
            Inc(Count);
        End;

        Comp := Iterator.NextPCBObject;
    End;

    // Make the mount list
    Comp := Iterator.FirstPCBObject;
    Count := 0;
    While (Comp <> Nil) Do
    Begin

        CompNum := 0;

        For I := 0 to PartsIndex.Count - 1 do
        Begin
            If (CompareStr(PartsIndex[I], RemoveIllegalChars(Comp.SourceLibReference)) = 0) then CompNum := I + 1;
        End;
        If CompNum > 0 Then
        Begin
             Inc(Count);
             MountList.Add('     <Mount No="' + IntToStr(Count) + '" X="'+ FloatToStr(CoordToMMs(Comp.X - Board.XOrigin))
                + '" Y="' + FloatToStr(CoordToMMs(Comp.Y - Board.YOrigin))
                + '" R="' + FloatToStr(Comp.Rotation) + '" FidRef="-0000001" BadRef="-0000001" CompRef="' + IntToStr(CompNum - 1) + '" ID="' + IntToStr(Count - 1) + '" Exist="1" Comment="'
                + Comp.SourceDesignator + '" CompNum="' + IntToStr(CompNum) + '"/>');

             MountVarList.Add('          <Mount No="' + IntToStr(Count) + '" PickGroup="-1" PickOrder="1" MountGroup="-1" MountOrder="' + IntToStr(Count) + '" Machine="6535" Head="2" Nozzle="2" Exec="0" PosMountRef="' + IntToStr(Count - 1) + '" SetupFeederRef="' + IntToStr(CompNum) + '" PosBlockRef="-0000001" OptFlag="0"/>');
        End;
        Comp := Iterator.NextPCBObject;
    End;

    Board.BoardIterator_Destroy(Iterator);

End;
{..............................................................................}

{..............................................................................}
Procedure IterateBoardArray;
Var
    Board         : IPCB_Board;
    Comp          : IPCB_ComponentBody;
    Iterator      : IPCB_BoardIterator;
    BdArray       : IPCB_EmbeddedBoard;

    Rpt           : TStringList;

    FileName      : TPCBString;
    Document      : IServerDocument;
    Count         : Integer;
    I             : Integer;
    J             : Integer;
    BR            : TCoordRect;
    OtherUnit     : TUnit;
    Edge_Str      : TPCBString;
    Area_Str      : TPCBString;
    Rows, R       : Integer;
    Columns, C    : Integer;
    X_Pos         : Float;
    Y_Pos         : Float;
    Col_Spc       : Float;
    Row_Spc       : Float;
    X_offset      : Float;
    Y_offset      : Float;
    
Begin
    // Retrieve the current board
    Board := PCBServer.GetCurrentPCBBoard;
    If Board = Nil Then Exit;

    // Get the panel dimensions
    Query_OutlinePerimeter(BR, OtherUnit, Edge_Str, Area_Str);

    // Create the iterator that will look for board arrays
    Iterator := Board.BoardIterator_Create;
    Iterator.AddFilter_ObjectSet(MkSet(eEmbeddedBoardObject));
    Iterator.AddFilter_LayerSet(AllLayers);
    Iterator.AddFilter_Method(eProcessAll);

    // Search for arrays and get their info
    Count := 0;
    Rpt := TStringList.Create;
    BdArray := Iterator.FirstPCBObject;

    Rpt.Add('<?xml version="1.0"?>');
    Rpt.Add('<PcbDataFile>');
    Rpt.Add('<LastEditing Date="2021/07/21" Time="11:17:54"/>');
    Rpt.Add('<Version No="1"/>');
    Rpt.Add('<PatternDatas>');
    Rpt.Add('    <PcbData Comment="' + RemovePath(ChangeFileExt(Board.FileName,'')) + '">');
    Rpt.Add('        <Board>');
    Rpt.Add('            <Origin X="xxx" Y="yyy" R="0.000"/>');
    Rpt.Add('            <Size X="' + FloatToStr(CoordToMMs(BR.Right - BR.Left)) + '" Y="' + FloatToStr(CoordToMMs(BR.Top - BR.Bottom)) + '" Z="1.600"/>');
    Rpt.Add('            <Block Number="0"/>');
    Rpt.Add('        </Board>');
    Rpt.Add('        <Fiducial>');
    Rpt.Add('            <Use Pcb="1" Blk="1" Local="0"/>');
    Rpt.Add('            <Pcb X1="25.481" Y1="2.055" Mark1Ref="00000000" X2="259.539" Y2="-244.972" Mark2Ref="00000000"/>');
    Rpt.Add('            <Blk X1="25.470" Y1="2.080" Mark1Ref="00000000" X2="81.042" Y2="-39.440" Mark2Ref="00000000"/>');
    Rpt.Add('        </Fiducial>');
    Rpt.Add('        <Badmark>');
    Rpt.Add('            <Use Pcb="1" Blk="1" Local="1"/>');
    Rpt.Add('            <Pcb X="0.000" Y="0.000" MarkRef="00000000"/>');
    Rpt.Add('            <Blk X="0.000" Y="0.000" MarkRef="00000000"/>');
    Rpt.Add('        </Badmark>');
    Rpt.Add('        <Production>');
    Rpt.Add('            <PcbFix Way="1"/>');
    Rpt.Add('            <Conveyor PreFixTimer="0.5" TransHeight="15" ConvTimer="0.0" YSpeed="3" TransType="0" MotorSpeed="0" PartsHeight="0.0"/>');
    Rpt.Add('            <Mount Exec="0" VacCheck="0" Alignment="0" CoPlanarity="0" PrePick="1" RetrySeq="1" TrayPre="0"/>');
    Rpt.Add('            <Dispense PreDisp="1" DotDisp="1" DotCheck="0" Refresh="0"/>');
    Rpt.Add('            <Exec Bin="00000101000000000000000000002C"/>');
    Rpt.Add('            <Trns Bin="010000010000000000050F03000000"/>');
    Rpt.Add('            <Cond Bin="000000000000000101010100010100"/>');
    Rpt.Add('            <Special Function="0" BackUpType="0"/>');
    Rpt.Add('        </Production>');
    Rpt.Add('        <Nozzle>');
    Rpt.Add('            <Variation Machine="6535">');
    Rpt.Add('                <Head A1="1" A2="0" A3="-1" A4="-1" A5="-1" A6="-1" A7="-1" A8="-1" A9="-1" A10="-1" A11="-1" A12="-1" A13="-1" A14="-1" A15="-1" A16="-1" B1="-1" B2="-1" B3="-1" B4="-1" B5="-1" B6="-1" B7="-1" B8="-1" B9="-1" B10="-1" B11="-1" B12="-1" B13="-1" B14="-1" B15="-1" B16="-1"/>');
    Rpt.Add('            </Variation>');
    Rpt.Add('        </Nozzle>');
    Rpt.Add('        <Special>');
    Rpt.Add('        </Special>');
    Rpt.Add('    </PcbData>');
    Rpt.Add('<Mounts>');

    IterateComponents(BdArray.DocumentPath);
    For R := 0 to MountList.Count - 1 do
    Begin
      Rpt.Add(MountList[R]);
    End;
    Rpt.Add('</Mounts>');
    Rpt.Add('<BlockRepeats>');
    
    While (BdArray <> Nil) Do
    Begin
        
        Rows := (BdArray.RowCount - 1);
        Columns := (BdArray.ColCount - 1);
        Row_Spc := CoordToMMs(BdArray.RowSpacing);
        Col_Spc := CoordToMMs(BdArray.ColSpacing);
        X_offset := CoordToMMs(BdArray.XLocation - Board.XOrigin) + 5.0;
        Y_offset := CoordToMMs(BdArray.YLocation - Board.YOrigin) - 5.0;
        // Assume bottom left corner of panel is CAD origin, bottom right corner is PnP origin.
        Rpt[7] := '            <Origin X="-' + FloatToStr(CoordToMMs(BR.Right - BR.Left) - X_offset) + '" Y="' + FloatToStr((Rows * Row_Spc) + Y_offset) + '" R="0.000"/>';
        For R := 0 to Rows  Do
        Begin
            For C := 0 to Columns Do
                Begin
                    Inc(Count);
                    Rpt.Add('   <Repeat No="' + IntToStr(Count) + '" X="' + FloatToStr(Col_Spc * C) + '" Y="' + FloatToStr(0 - (Row_Spc * R)) 
                    + '" R="0.000" Exec="0" ID="' + IntToStr(Count) + '" Exist="1" Comment="Board_' + IntToStr(Count) + '"/>');
                End;
        End;
        
        //Rpt.Add(' X Origin : ' + CoordUnitToString(Board.XOrigin, eMM));
        //Rpt.Add(' Y Origin : ' + CoordUnitToString(Board.YOrigin, eMM));

        BdArray := Iterator.NextPCBObject;
    End;
    
    Rpt.Add('</BlockRepeats>');
    Rpt.Add('<LocalFiducials>');
    Rpt.Add('</LocalFiducials>');
    Rpt.Add('<LocalBadmarks>');
    Rpt.Add('</LocalBadmarks>');
    Rpt.Add('<PreDispenses>');
    Rpt.Add('</PreDispenses>');
    Rpt.Add('<DotDispenses>');
    Rpt.Add('</DotDispenses>');
    Rpt.Add('</PatternDatas>');
    Rpt.Add('<OptimizedVariations>');
    Rpt.Add('    <MountVariation>');
    Rpt.Add('        <Optimized>');
    For R := 0 to MountList.Count - 1 do
    Begin
      Rpt.Add(MountVarList[R]);
    End;
    Rpt.Add('        </Optimized>');
    Rpt.Add('        <Set>');
    For R := 0 to FeederList.Count - 1 do
    Begin
      Rpt.Add(FeederList[R]);
    End;
    Rpt.Add('        </Set>');
    Rpt.Add('    </MountVariation>');
    Rpt.Add('    <PreDispensesVariation>');
    Rpt.Add('        <Optimized>');
    Rpt.Add('        </Optimized>');
    Rpt.Add('    </PreDispensesVariation>');
    Rpt.Add('    <DotDispensesVariation>');
    Rpt.Add('        <Optimized>');
    Rpt.Add('        </Optimized>');
    Rpt.Add('    </DotDispensesVariation>');
    Rpt.Add('</OptimizedVariations>');
    Rpt.Add('<Components>');
    
    For I := 0 to PartsList.Count - 1 do
    Begin
        Rpt.Add(PartsList[I]);
        Rpt.Add('       <Library Use="0"/>');
        Rpt.Add('       <CompNumList Machine="6535" IBMFNo="'+ IntToStr(I + 1) +'"/>');
        Rpt.Add('       <Feeder Package="0" Type="0" Pitch="0"/>');
        Rpt.Add('       <Shape Type="7">');
        Rpt.Add(SizeList[I]);
        Rpt.Add('           <Lead>');
        Rpt.Add('               <Common RulerOffset="3"/>');
        Rpt.Add('           </Lead>');
        Rpt.Add('           <AlignSize X="1.700" Y="3.300" Z="0.400"/>');
        Rpt.Add('       </Shape>');
        Rpt.Add('       <Vision Module="0">');
        Rpt.Add('           <Light Setting="0" AllLightOn="0"/>');
        Rpt.Add('           <Recognition AutoThreshold="1" LightLevel="3" Threshold="30" Torelance="30" SerachArea="0.800" InsideRecognition="0"/>');
        Rpt.Add('           <Datum Angle="0"/>');
        Rpt.Add('           <MultiMACS Use="0"/>');
        Rpt.Add('           <Co-Planarity CompIntensity="0" Level="0" Threshold="0" Ruler="0"/>');
        Rpt.Add('           <DDD DDDThreshold="70" DDDLightCoax="4" DDDLightMain="3" DDDLightSide="4" DDDBrightArea="25" DDDHeightCheck="0.000" DDDExecPass="3"/>');
        Rpt.Add('           <Laser ErrDetect="0"/>');
        Rpt.Add('       </Vision>');
        Rpt.Add('       <PickMount>');
        Rpt.Add('           <Nozzle Required="3"/>');
        Rpt.Add('           <Vacuum Check="1"/>');
        Rpt.Add('           <Pick Timer="0" Height="0.000" Action="0" Speed="0" Angle="0" Level="30" CorrectPos="0" SingleDir="0" Down="0" Up="0" SecondSrvDown="0" SecondSrvUp="0"/>');
        Rpt.Add('           <Mount Timer="0" Height="0.500" Action="0" Speed="0" XYSpeed="0" VacuumLevel="30" SingleDir="0" Down="0" Up="0" SecondSrvDown="0" SecondSrvUp="0"/>');
        Rpt.Add('           <Dump Point="0" Retry="3"/>');
        Rpt.Add('           <ConvY Speed="3"/>');
        Rpt.Add('           <Count OutStop="0"/>');
        Rpt.Add('       </PickMount>');
        Rpt.Add('       <Dispense>');
        Rpt.Add('           <Nozzle Size="3" Shot="0" Unit="0"/>');
        Rpt.Add('           <Position>');
        Rpt.Add('               <Ref X="0.000" Y="0.000" R="0.000"/>');
        Rpt.Add('               <DotExt X="0.000" Y="0.000"/>');
        Rpt.Add('               <Amount X="1" Y="1"/>');
        Rpt.Add('           </Position>');
        Rpt.Add('       </Dispense>');
        Rpt.Add('   </Component>');     
    End;
    
    
    Rpt.Add('</Components>');
    Rpt.Add('<Marks>'); // fiducial marks, just default to one 1mm circular mark
    Rpt.Add('    <Mark No="1" Name="Cir_1.0_NORMAL" Comment="REFLECT" ID="00000000" DatabaseNo="156">');
    Rpt.Add('        <Library Use="0"/>');
    Rpt.Add('        <MarkNumList Machine="6535" IBMFNo="1"/>');
    Rpt.Add('        <Common Type="1"/>');
    Rpt.Add('        <Shape Shape="0" Reflect="1" Algorithm="0">');
    Rpt.Add('            <OutSize X="1.000"/>');
    Rpt.Add('        </Shape>');
    Rpt.Add('        <Vision>');
    Rpt.Add('            <Vison Sequence="0" Mode="0"/>');
    Rpt.Add('            <Recognition Threshold="119" Tolerance="30" SearchAreaX="3.00" SearchAreaY="3.00" SearchAreaOffsetX="0.000" SearchAreaOffsetY="0.000" MaxDefPosXY="0" MaxDefPosR="0.00" Height="0.000"/>');
    Rpt.Add('            <Filter Inner="0" Outer="0"/>');
    Rpt.Add('            <Lighting Outer="2" Inner="2" Drop="2" IROuter="0" IRInner="0"/>');
    Rpt.Add('            <DispenseDot DotRecognition="0" StandardArea="1.770" AreaTol="30" ShapeProportion="0.000" ShapeTorelance="0" AreaTorelanceToCheck="0" ShapeTorelanceToCheck="0" MinSample="0" LightType="0"/>');
    Rpt.Add('        </Vision>');
    Rpt.Add('    </Mark>');
    Rpt.Add('</Marks>');
    Rpt.Add('</PcbDataFile>');

    Board.BoardIterator_Destroy(Iterator);

    // Display the Component Bodies report
    FileName := ChangeFileExt(Board.FileName,'_PnP.xml');
    Rpt.SaveToFile(Filename);
    Rpt.Free;

    Document  := Client.OpenDocument('Text', FileName);
    If Document <> Nil Then
        Client.ShowDocument(Document);
End;
{..............................................................................}

procedure FillList;
    Var
    Board         : IPCB_Board;
    Comp          : IPCB_Component;
    Iterator      : IPCB_BoardIterator;
    BdArray       : IPCB_EmbeddedBoard;

    I             : Integer;
    J             : Integer;
    K             : Integer;
    Added         : Integer;
    S             : TString;
    CompNum       : Integer;
    Count         : Integer;
    IncludeList   : TStringList;

Begin
    // Retrieve the current board
    Board := PCBServer.GetCurrentPCBBoard;
    If Board = Nil Then Exit;

    // Create the iterator that will look for board arrays
    Iterator := Board.BoardIterator_Create;
    Iterator.AddFilter_ObjectSet(MkSet(eEmbeddedBoardObject));
    Iterator.AddFilter_LayerSet(AllLayers);
    Iterator.AddFilter_Method(eProcessAll);
    BdArray := Iterator.FirstPCBObject;

    // Open the board from the embedded array
    Board := PCBServer.GetPCBBoardByPath(BdArray.DocumentPath);
    If Board = Nil Then Exit;

    // Create the iterator that will look for components only
    Iterator := Board.BoardIterator_Create;
    Iterator.AddFilter_ObjectSet(MkSet(eComponentObject));
    Iterator.AddFilter_LayerSet(MkSet(SelectedLayer));
    Iterator.AddFilter_Method(eProcessAll);

    IncludeList := TStringList.Create;

    Comp := Iterator.FirstPCBObject;

    While (Comp <> Nil) Do
    Begin
        S := Comp.SourceLibReference;
        // check if this component is already in the parts list, then add it if not
        Added := 0;
        for K := 0 to IncludeList.Count - 1 do
        Begin
             if (CompareStr(IncludeList[K], S) = 0) then Added := 1;
        End;

        if (Added = 0) then
        Begin
            IncludeList.Add(S);
            CheckListBox1.Items.Add(S);
            Inc(Count);
        End;
        Comp := Iterator.NextPCBObject;
    End;

    Board.BoardIterator_Destroy(Iterator);

End;

procedure TForm1.XPRadioButton1Click(Sender: TObject);
begin
     CheckListBox1.Items.Clear;
     SelectedLayer := eTopLayer;
     FillList;
end;

procedure TForm1.XPRadioButton2Click(Sender: TObject);
begin
     CheckListBox1.Items.Clear;
     SelectedLayer := eBottomLayer;
     FillList;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
     IterateBoardArray;
end;

procedure TForm1.Form1Create(Sender: TObject);
var
   I : Integer;
begin
     SelectedLayer := eTopLayer;
     FillList;
     For I := 0 to CheckListBox1.Items.Count - 1 Do CheckListBox1.Checked[I] := True;
end;

