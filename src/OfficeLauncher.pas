//
// StartBox
//
// Autor: Michael Tesch, Bredstedt
//
// Anfang: 08.06.2026
// Ende:   12.06.2026
//

unit OfficeLauncher;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI,
  Winapi.ShlObj, Vcl.Dialogs, System.Math,
  System.SysUtils, System.Classes, System.IniFiles,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.ExtCtrls, Vcl.ImgList, Vcl.Menus,
  System.Generics.Collections, System.ImageList;

type
  TForm1 = class(TForm)
    ImageList1: TImageList;
    TrayIcon1: TTrayIcon;
    PopupMenu1: TPopupMenu;
    MenuExit: TMenuItem;
    MenuHelp: TMenuItem;
    MinimizeJN: TMenuItem;
    Restore: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;

    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TrayIcon1Click(Sender: TObject);
    procedure MenuExitClick(Sender: TObject);

    procedure ButtonClick(Sender: TObject);
    procedure ButtonMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ButtonMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure ButtonMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure MenuHelpClick(Sender: TObject);
    procedure MinimizeJNClick(Sender: TObject);
    procedure RestoreClick(Sender: TObject);

  private
    Ini: TIniFile;
    Buttons: TList<TBitBtn>;
    ForceClose: Boolean;

    ButtonHeight: Integer;
    ButtonSpacing: Integer;

    IsDragging: Boolean;
    DragButton: TBitBtn;
    DragOffsetY: Integer;
    DragOffsetX: Integer;
    DragTotalMove: Integer;
    DragDelayPassed: Boolean;
    DragStartTime: Cardinal;
    DragDelay: Cardinal;
    DragStartX, DragStartY: Integer;
    LastSwapButton: TBitBtn;

    OfficeEdition: string;

    PopupButton: TPopupMenu;
    MenuDeleteButton: TMenuItem;
    MenuRenameButton: TMenuItem;
    ButtonToDelete: TBitBtn;

    procedure WMSysCommand(var Msg: TWMSysCommand); message WM_SYSCOMMAND;
    procedure WMDropFiles(var Msg: TWMDropFiles); message WM_DROPFILES;

    procedure MinimizeToTray;

    function DetectOfficeEdition: string;
    function FindOfficePath: string;

    procedure AddButton(const Caption, ExeName: string);
    procedure LoadOfficeButtons;
    procedure LoadButtonsFromIni;
    procedure SaveButtonsToIni;

    procedure LoadFormPosition;
    procedure SaveFormPosition;
    function  ButtonExistsForExe(const Exe: string): Boolean;

    procedure RecalculateButtonWidths;
    procedure RecalculateButtonPositions;
    function  CalculateOptimalButtonWidth: Integer;

    function ButtonsOverlap(A, B: TBitBtn): Boolean;
    procedure SwapButtons(A, B: TBitBtn);

    procedure ButtonContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure MenuDeleteButtonClick(Sender: TObject);
    procedure MenuRenameButtonClick(Sender: TObject);

    function ResolveLnk(const LnkFile: string): string;

  public
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses
  Winapi.ActiveX, System.Win.ComObj, Vcl.Graphics;

const
  OfficeApps: array[0..5] of record
    Exe: string;
    Caption: string;
  end = (
    (Exe: 'winword.exe'; Caption: 'Word'),
    (Exe: 'excel.exe';   Caption: 'Excel'),
    (Exe: 'powerpnt.exe';Caption: 'PowerPoint'),
    (Exe: 'outlook.exe'; Caption: 'Outlook'),
    (Exe: 'onenote.exe'; Caption: 'OneNote'),
    (Exe: 'msaccess.exe';Caption: 'Access')
  );
  ButtonWidth   = 200;
  ButtonHeight  = 40;
  ButtonSpacing = 10;

procedure TForm1.WMSysCommand(var Msg: TWMSysCommand);
begin
  if (Msg.CmdType and $FFF0) = SC_MINIMIZE then
  begin
    MinimizeToTray;
    Exit;
  end;
  inherited;
end;

function ExtractDomainFromURL(const URL: string): string;
var
  S: string;
  P: Integer;
begin
  Result := URL;

  // Protokoll entfernen
  S := URL;
  if S.StartsWith('http://') then
    Delete(S, 1, 7)
  else if S.StartsWith('https://') then
    Delete(S, 1, 8);

  // Pfad abschneiden
  P := Pos('/', S);
  if P > 0 then
    S := Copy(S, 1, P - 1);

  // www. entfernen
  if S.StartsWith('www.') then
    Delete(S, 1, 4);

  Result := S;
end;

function ReadUrlFromInternetShortcut(const FileName: string): string;
var
  SL: TStringList;
  I: Integer;
begin
  Result := '';
  SL := TStringList.Create;
  try
    SL.LoadFromFile(FileName);
    for I := 0 to SL.Count - 1 do
      if SL[I].StartsWith('URL=') then
        Exit(Copy(SL[I], 5, MaxInt));
  finally
    SL.Free;
  end;
end;

procedure TForm1.MinimizeJNClick(Sender: TObject);
begin
  MinimizeJN.Checked := not MinimizeJN.Checked;
end;

procedure TForm1.MinimizeToTray;
begin
  TrayIcon1.Visible := True;
  Hide;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if not ForceClose then
  begin
    CanClose := False;
    MinimizeToTray;
  end else
    CanClose := True;
end;

procedure TForm1.TrayIcon1Click(Sender: TObject);
begin
  TrayIcon1.Visible := False;
  Show;
  WindowState := wsNormal;
  BringToFront;
end;

procedure TForm1.MenuExitClick(Sender: TObject);
begin
  ForceClose := True;
  Close;
end;

// MessageDlg darstellen
function MessageDlgBottomRight(const Msg: string; DlgType: TMsgDlgType;
  Buttons: TMsgDlgButtons): Integer;
var
  WorkArea: TRect;
begin
  with CreateMessageDialog(Msg, DlgType, Buttons) do
  try
    // Arbeitsbereich ohne Taskleiste holen
    SystemParametersInfo(SPI_GETWORKAREA, 0, @WorkArea, 0);

    // Rechts unten positionieren
    Left := WorkArea.Right - Width - 20;   // 20px Abstand vom Rand
    Top  := WorkArea.Bottom - Height - 20; // 20px über Taskleiste

    Result := ShowModal;
  finally
    Free;
  end;
end;

procedure TForm1.MenuHelpClick(Sender: TObject);
begin
  MessageDlgBottomRight
    ('StartBox Version 1.1' + #13 + #13 +
     'Copyright © 2026 by FreePDF64@outlook.com' + #13 +
     'Website -> https://github.com/FreePDF64/StartBox' + #13 + #13 +
     'StartBox darf sowohl im privaten als auch im kommerziellen' + #13 +
     'Umfeld ohne Bezahlung eingesetzt werden ("Freeware")!' + #13 +
     'Der Autor übernimmt keinerlei Haftung für Fehler, die direkt' + #13 +
     'oder indirekt aus der Benutzung dieser Software entstehen.' + #13 + #13 +
     'ToDo:' + #13 +
     '- Programme/Dateien/etc. werden per Drag&&Drop hinzugefügt' + #13 +
     '- Icons werden automatisch aus den Dateien extrahiert' + #13 +
     '- Rechtsklick auf Button: Umbenennen oder Löschen' + #13 +
     '- Schließen (X) minimiert in den Tray' + #13 +
     '- Auf Wunsch NACH Klick auf Buttons in den Tray minimieren' + #13 +
     '- Beenden NUR über das Tray-Menü',
    mtInformation, [mbOk]);
end;

function TForm1.DetectOfficeEdition: string;
var P: string;
begin
  P := FindOfficePath.ToLower;
  if P.Contains('office16') then Exit('Office 2024');
  if P.Contains('office15') then Exit('Office 2016/2021/365');
  Result := 'Office';
end;

function TForm1.FindOfficePath: string;
const
  Paths: array[0..7] of string = (
    'C:\Program Files\Microsoft Office\root\Office24\',
    'C:\Program Files (x86)\Microsoft Office\root\Office24\',
    'C:\Program Files\Microsoft Office\root\Office16\',
    'C:\Program Files (x86)\Microsoft Office\root\Office16\',
    'C:\Program Files\Microsoft Office\root\Office15\',
    'C:\Program Files (x86)\Microsoft Office\root\Office15\',
    'C:\Program Files\Microsoft Office\Office24\',
    'C:\Program Files\Microsoft Office\Office16\'
  );
var P: string;
begin
  for P in Paths do
    if DirectoryExists(P) then Exit(P);
  Result := '';
end;

function ExtractShellIconToImageList(const FileName: string; ImageList: TImageList): Integer;
var
  SFI: SHFILEINFO;
  Icon: TIcon;
begin
  Result := -1;

  // Shell-Icon für die Datei holen (16x16)
  if SHGetFileInfo(PChar(FileName), 0, SFI, SizeOf(SFI),
     SHGFI_ICON or SHGFI_SMALLICON) = 0 then
    Exit;

  Icon := TIcon.Create;
  try
    Icon.Handle := SFI.hIcon;
    Result := ImageList.AddIcon(Icon);
  finally
    Icon.Free;
  end;

  // Icon-Handle freigeben
  DestroyIcon(SFI.hIcon);
end;

// Einspaltig, sonst auf zwei Spalten erweitern...
procedure TForm1.RecalculateButtonPositions;
const
  MarginX = 10;
  MarginY = 6;
  TopOffset = 10;
  LeftOffset = 10;
var
  I, Count: Integer;
  BtnW, BtnH: Integer;
  NeededHeight: Integer;
  MaxH, MaxW: Integer;
  Cols, Rows: Integer;
  BtnLeft, BtnTop: Integer;
begin
  Count := Buttons.Count;
  if Count = 0 then Exit;

  BtnW := Buttons[0].Width;
  BtnH := Buttons[0].Height;

  // Bildschirmgrenzen (ohne Taskleiste)
  MaxH := Screen.WorkAreaHeight;
  MaxW := Screen.WorkAreaWidth;

  // Höhe, wenn ALLE Buttons untereinander stehen
  NeededHeight := TopOffset + Count * (BtnH + MarginY) + 20;

  // Standard: 1 Spalte
  Cols := 1;

  // Wenn die Form zu hoch würde → auf 2 Spalten umschalten
  if NeededHeight > MaxH then
    Cols := 2;

  // Anzahl benötigter Zeilen
  Rows := (Count + Cols - 1) div Cols;

  // Formgröße anpassen
  ClientWidth  := Cols * (BtnW + MarginX) + LeftOffset * 2;
  ClientHeight := Rows * (BtnH + MarginY) + TopOffset * 2;

  // *** NEU: Wenn die Form rechts aus dem Monitor ragt → nach links schieben ***
  if Left + Width > Screen.WorkAreaLeft + MaxW then
    Left := (Screen.WorkAreaLeft + MaxW) - Width;

  // Buttons positionieren
  for I := 0 to Count - 1 do
  begin
    BtnLeft := LeftOffset + (I mod Cols) * (BtnW + MarginX);
    BtnTop  := TopOffset + (I div Cols) * (BtnH + MarginY);

    Buttons[I].SetBounds(BtnLeft, BtnTop, BtnW, BtnH);
  end;
end;

procedure TForm1.AddButton(const Caption, ExeName: string);
var
  Btn: TBitBtn;
  IconIndex: Integer;
  Cap: string;
begin
  Btn := TBitBtn.Create(Self);
  Btn.Parent := Self;

  // URL?
  if ExeName.StartsWith('http://') or ExeName.StartsWith('https://') then
  begin
    Cap := ExtractDomainFromURL(ExeName);
    Btn.Glyph      := NIL;
    Btn.Images     := ImageList1;
    Btn.ImageIndex := 0;
  end
  else
  begin
    Cap := Caption;

    IconIndex := ExtractShellIconToImageList(ExeName, ImageList1);
    Btn.Glyph := nil;
    Btn.Images := ImageList1;
    Btn.ImageIndex := IconIndex;
  end;

  Btn.Caption := Cap;
  Btn.Hint := ExeName;
  Btn.ShowHint := True;
  Btn.Font.Name := 'Segoe UI';
  Btn.Font.Size := 9;
  Btn.Height := ButtonHeight;

  Btn.PopupMenu := PopupButton;
  Btn.OnContextPopup := ButtonContextPopup;

  Btn.OnMouseDown := ButtonMouseDown;
  Btn.OnMouseMove := ButtonMouseMove;
  Btn.OnMouseUp := ButtonMouseUp;
  Btn.OnClick := ButtonClick;

  Buttons.Add(Btn);

  // Breite neu berechnen (deine Funktion)
  RecalculateButtonWidths;
  RecalculateButtonPositions;  // NEUE Version mit Spaltenlogik
end;

procedure TForm1.LoadOfficeButtons;
var
  Base: string;

  procedure AddIfExists(const FileName, Caption: string);
  var
    FullPath: string;
  begin
    FullPath := Base + FileName;

    if FileExists(FullPath) then
    begin
      // Nur hinzufügen, wenn nicht bereits in der INI vorhanden
      if not ButtonExistsForExe(FullPath) then
        AddButton(Caption + ' (' + OfficeEdition + ')', FullPath);
    end;
  end;

begin
  // Office-Pfad ermitteln
  Base := FindOfficePath;
  OfficeEdition := DetectOfficeEdition;

  AddIfExists('WINWORD.EXE', 'Word');
  AddIfExists('EXCEL.EXE', 'Excel');
  AddIfExists('POWERPNT.EXE', 'PowerPoint');
  AddIfExists('OUTLOOK.EXE', 'Outlook');
  AddIfExists('ONENOTE.EXE', 'OneNote');
  AddIfExists('MSACCESS.EXE', 'Access');
  AddIfExists('MSPUB.EXE', 'Publisher');
end;

procedure TForm1.LoadButtonsFromIni;
var
  I, Count: Integer;
  Exe, Cap: string;
begin
  Count := Ini.ReadInteger('Buttons', 'Count', 0);

  for I := 0 to Count - 1 do
  begin
    Exe := Ini.ReadString('Buttons', 'Btn' + I.ToString + '_Exe', '');
    Cap := Ini.ReadString('Buttons', 'Btn' + I.ToString + '_Caption', '');

    if Exe = '' then Continue;

    // URL?
    if Exe.StartsWith('http://') or Exe.StartsWith('https://') then
    begin
      if Cap = '' then
        Cap := ExtractDomainFromURL(Exe);

      AddButton(Cap, Exe);
      Continue;
    end;

    // EXE?
    if FileExists(Exe) then
    begin
      if Cap = '' then
        Cap := ChangeFileExt(ExtractFileName(Exe), '');

      AddButton(Cap, Exe);
      Continue;
    end;
     // Fallback (z. B. Datei nicht mehr vorhanden)
     AddButton(Cap, Exe);
  end;
end;

procedure TForm1.SaveButtonsToIni;
var
  I: Integer;
begin
  Ini.EraseSection('Buttons');
  Ini.WriteInteger('Buttons', 'Count', Buttons.Count);
  for I := 0 to Buttons.Count - 1 do
  begin
    Ini.WriteString('Buttons', 'Btn' + I.ToString + '_Exe', Buttons[I].Hint);
    Ini.WriteString('Buttons', 'Btn' + I.ToString + '_Caption', Buttons[I].Caption);
  end;
end;

procedure TForm1.LoadFormPosition;
begin
  Left := Ini.ReadInteger('Form', 'Left', Left);
  Top := Ini.ReadInteger('Form', 'Top', Top);
  MinimizeJN.Checked := Ini.ReadBool('Form', 'Minimize', MinimizeJN.Checked);
end;

procedure TForm1.SaveFormPosition;
begin
  Ini.WriteInteger('Form', 'Left', Left);
  Ini.WriteInteger('Form', 'Top', Top);
  Ini.WriteBool('Form', 'Minimize', MinimizeJN.Checked);
end;

function TForm1.CalculateOptimalButtonWidth: Integer;
var
  I, W: Integer;
begin
  Result := 140; // Mindestbreite etwas größer

  Canvas.Font.Assign(Buttons[0].Font);

  for I := 0 to Buttons.Count - 1 do
  begin
    // 40px für Icon + Padding, 20px extra für mehr Luft
    W := Canvas.TextWidth(Buttons[I].Caption) + 90;
    if W > Result then
      Result := W;
  end;
end;

procedure TForm1.RecalculateButtonWidths;
var
  I: Integer;
  NewWidth: Integer;
begin
  if Buttons.Count = 0 then Exit;

  NewWidth := CalculateOptimalButtonWidth;

  for I := 0 to Buttons.Count - 1 do
    Buttons[I].Width := NewWidth;
end;

function TForm1.ButtonExistsForExe(const Exe: string): Boolean;
var
  Btn: TBitBtn;
begin
  Result := False;
  for Btn in Buttons do
    if SameText(Btn.Hint, Exe) then
      Exit(True);
end;

procedure TForm1.FormCreate(Sender: TObject);
var IniPath: string;
begin
  CoInitialize(nil);

  Application.HintPause := 700;       // Verzögerung bis Tooltip erscheint
  Application.HintHidePause := 5000;  // Wie lange der Tooltip sichtbar bleibt
  DragDelay := 300;                   // Verzögerung von Drag&Drop in Millisekunden

  IniPath := ChangeFileExt(Application.ExeName, '.ini');
  Ini := TIniFile.Create(IniPath);
  Buttons := TList<TBitBtn>.Create;
  ButtonHeight := 50;
  ButtonSpacing := 10;
  ForceClose := False;

  ImageList1.Width  := 16;
  ImageList1.Height := 16;
  ImageList1.ColorDepth := cd32Bit;
  ImageList1.Masked := True;

  PopupButton := TPopupMenu.Create(Self);
  MenuDeleteButton := TMenuItem.Create(Self);
  MenuDeleteButton.Caption := 'Löschen';
  MenuDeleteButton.OnClick := MenuDeleteButtonClick;
  PopupButton.Items.Add(MenuDeleteButton);
  MenuRenameButton := TMenuItem.Create(Self);
  MenuRenameButton.Caption := 'Umbenennen';
  MenuRenameButton.OnClick := MenuRenameButtonClick;
  PopupButton.Items.Insert(0, MenuRenameButton);

  DragAcceptFiles(Handle, True);

  ShowHint := True;
  Hint := 'Info über StartBox -> Rechtsklick auf Symbol im SysTray';

  LoadFormPosition;
  OfficeEdition := DetectOfficeEdition;

  LoadButtonsFromIni;     // Benutzerdefinierte Buttons
  LoadOfficeButtons;      // Office automatisch hinzufügen

  if Buttons.Count = 0 then
    MessageDlg('Bitte Anwendungen/Dateien/etc. hinzufügen per Drag&Drop.', mtInformation, [mbOK], 0);

  RecalculateButtonWidths;
  RecalculateButtonPositions;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  SaveFormPosition;
  SaveButtonsToIni;
  Ini.Free;
  Buttons.Free;
  CoUninitialize;
end;

procedure OpenWithDialog(const FileName: string);
var
  sei: TShellExecuteInfo;
begin
  ZeroMemory(@sei, SizeOf(sei));
  sei.cbSize := SizeOf(sei);
  sei.fMask := SEE_MASK_INVOKEIDLIST;
  sei.lpVerb := 'openas';  // <<< moderner Windows 11 Dialog
  sei.lpFile := PChar(FileName);
  sei.nShow := SW_SHOWNORMAL;

  ShellExecuteEx(@sei);
end;

procedure TForm1.ButtonClick(Sender: TObject);
var
  Target: string;
  ExecResult: HINST;
begin
  if IsDragging then Exit;

  Target := TBitBtn(Sender).Hint;

  // Erst normal versuchen
  ExecResult := ShellExecute(0, 'open', PChar(Target), nil, nil, SW_SHOWNORMAL);

  // Wenn nicht startbar → Windows 11 „Öffnen mit…“-Dialog
  if ExecResult <= 32 then
  begin
    OpenWithDialog(Target);
    Exit;
  end;

  if MinimizeJN.Checked then
    MinimizeToTray;
end;

function TForm1.ButtonsOverlap(A, B: TBitBtn): Boolean;
begin
  Result := (A.Top < B.Top + B.Height) and (A.Top + A.Height > B.Top);
end;

procedure TForm1.SwapButtons(A, B: TBitBtn);
var IA, IB: Integer; Tmp: TBitBtn;
begin
  IA := Buttons.IndexOf(A);
  IB := Buttons.IndexOf(B);
  Tmp := Buttons[IA];
  Buttons[IA] := Buttons[IB];
  Buttons[IB] := Tmp;
end;

procedure TForm1.ButtonMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button <> mbLeft then Exit;

  DragButton := TBitBtn(Sender);

  // Startwerte setzen
  IsDragging := False;
  DragDelayPassed := False;

  DragStartTime := GetTickCount;
  DragStartX := X;
  DragStartY := Y;

  // Offset für visuelles Verschieben
  DragOffsetX := X;
  DragOffsetY := Y;

  // Für Click-Blockierung
  DragTotalMove := 0;

  LastSwapButton := nil;
end;

procedure TForm1.ButtonMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var
  Btn: TBitBtn; I, dx, dy: Integer;
begin
  // Mausbewegung prüfen
  dx := Abs(X - DragStartX);
  dy := Abs(Y - DragStartY);

  // Bewegungsschwelle (z. B. 4 Pixel)
  if (dx < 4) and (dy < 4) then
    Exit;

  // Zeitverzögerung prüfen
  if not DragDelayPassed then
  begin
    if GetTickCount - DragStartTime < DragDelay then
      Exit;

    DragDelayPassed := True;
  end;

  // Jetzt erst Drag starten
  IsDragging := True;
  if not Assigned(DragButton) then
    Exit;

  if IsDragging then
  begin
    DragTotalMove := DragTotalMove + Abs(Y - DragOffsetY);
    DragButton.Top := DragButton.Top + (Y - DragOffsetY);
    for I := 0 to Buttons.Count - 1 do
    begin
      Btn := Buttons[I];
      if (Btn <> DragButton) and ButtonsOverlap(DragButton, Btn) then
      begin
        if Btn <> LastSwapButton then
        begin
          SwapButtons(DragButton, Btn);
          LastSwapButton := Btn;
        end;
        Exit;
      end;
    end;
    LastSwapButton := nil;
  end;
end;

procedure TForm1.ButtonMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  IsDragging := False;
  DragButton := nil;
  LastSwapButton := nil;
end;

procedure TForm1.ButtonContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  ButtonToDelete := TBitBtn(Sender);
  Handled := False;
end;

procedure KeepFormOnScreen(F: TForm);
var
  R: TRect;
begin
  // Arbeitsbereich des Monitors holen (ohne Taskleiste)
  SystemParametersInfo(SPI_GETWORKAREA, 0, @R, 0);

  // Links korrigieren
  if F.Left < R.Left then
    F.Left := R.Left;

  // Rechts korrigieren
  if F.Left + F.Width > R.Right then
    F.Left := R.Right - F.Width;

  // Oben korrigieren
  if F.Top < R.Top then
    F.Top := R.Top;

  // Unten korrigieren
  if F.Top + F.Height > R.Bottom then
    F.Top := R.Bottom - F.Height;
end;

procedure TForm1.MenuDeleteButtonClick(Sender: TObject);
var
  dlg: TForm;
  res: Integer;
  p: TPoint;
  MsgText: string;
begin
  if ButtonToDelete = nil then Exit;

  // Text dynamisch erzeugen
  MsgText := 'Button löschen: "' + ButtonToDelete.Caption + '"';

  // Dialog erzeugen
  dlg := CreateMessageDialog(MsgText, mtConfirmation, [mbYes, mbNo]);

  try
    // Position des Buttons in Bildschirmkoordinaten
    p := ButtonToDelete.ClientToScreen(Point(0, 0));

    dlg.Position := poDesigned;
    // X = gleiche linke Position wie der Button
    dlg.Left := p.X;
    // Y = unterhalb des Buttons
    dlg.Top := p.Y + ButtonToDelete.Height + 10;
    // Dialog anzeigen
    KeepFormOnScreen(dlg);
    res := dlg.ShowModal;

    if res <> mrYes then
      Exit;

    // Button löschen
    Buttons.Remove(ButtonToDelete);
    ButtonToDelete.Free;
    ButtonToDelete := nil;

    RecalculateButtonWidths;
    RecalculateButtonPositions;
    SaveButtonsToIni;

  finally
    dlg.Free;
  end;
end;

procedure TForm1.MenuRenameButtonClick(Sender: TObject);
var
  dlg: TForm;
  edt: TEdit;
  lbl: TLabel;
  btnOK, btnCancel: TButton;
  p: TPoint;
  res: Integer;
begin
  if ButtonToDelete = nil then Exit;

  // Dialog erzeugen
  dlg := TForm.Create(Self);
  try
    dlg.BorderStyle := bsDialog;
    dlg.Caption := 'Umbenennen';
    dlg.Position := poDesigned;
    dlg.ClientWidth := 260;
    dlg.ClientHeight := 120;

    // Label
    lbl := TLabel.Create(dlg);
    lbl.Parent := dlg;
    lbl.Caption := 'Neuer Name:';
    lbl.Left := 10;
    lbl.Top := 10;

    // Eingabefeld
    edt := TEdit.Create(dlg);
    edt.Parent := dlg;
    edt.Left := 10;
    edt.Top := 30;
    edt.Width := dlg.ClientWidth - 20;
    edt.Text := ButtonToDelete.Caption;

    // OK-Button
    btnOK := TButton.Create(dlg);
    btnOK.Parent := dlg;
    btnOK.Caption := 'OK';
    btnOK.ModalResult := mrOk;
    btnOK.Left := dlg.ClientWidth - 220;
    btnOK.Top := 80;
    btnOK.Width := 100;

    // Abbrechen-Button
    btnCancel := TButton.Create(dlg);
    btnCancel.Parent := dlg;
    btnCancel.Caption := 'Abbrechen';
    btnCancel.ModalResult := mrCancel;
    btnCancel.Left := dlg.ClientWidth - 110;
    btnCancel.Top := 80;
    btnCancel.Width := 100;

    // Position unterhalb des Buttons
    p := ButtonToDelete.ClientToScreen(Point(0, 0));
    dlg.Left := p.X;
    dlg.Top := p.Y + ButtonToDelete.Height + 10;

    // Dialog anzeigen
    KeepFormOnScreen(dlg);
    res := dlg.ShowModal;

    if res <> mrOk then
      Exit;

    if Trim(edt.Text) = '' then
      Exit;

    // Neuen Namen setzen
    ButtonToDelete.Caption := edt.Text;

    // Layout aktualisieren
    RecalculateButtonWidths;
    RecalculateButtonPositions;

    // Speichern
    SaveButtonsToIni;
  finally
    dlg.Free;
  end;
end;

function TForm1.ResolveLnk(const LnkFile: string): string;
var
  ShellLink: IShellLink;
  PersistFile: IPersistFile;
  Target: array[0..MAX_PATH] of Char;
  FindData: TWin32FindData;
begin
  Result := '';

  if not FileExists(LnkFile) then
    Exit;

  // COM-Objekt erzeugen
  ShellLink := CreateComObject(CLSID_ShellLink) as IShellLink;
  PersistFile := ShellLink as IPersistFile;

  // .lnk laden
  if PersistFile.Load(PWideChar(WideString(LnkFile)), STGM_READ) = S_OK then
  begin
    // Zielpfad extrahieren
    if ShellLink.GetPath(Target, MAX_PATH, FindData, SLGP_UNCPRIORITY) = S_OK then
      Result := Target;
  end;
end;

procedure TForm1.RestoreClick(Sender: TObject);
begin
  TrayIcon1.Visible := False;
  Show;
  WindowState := wsNormal;
  BringToFront;
end;

procedure TForm1.WMDropFiles(var Msg: TWMDropFiles);
var
  Count: Integer;
  FileName: array[0..MAX_PATH] of Char;
  Dropped, URL: string;
begin
  Count := DragQueryFile(Msg.Drop, $FFFFFFFF, nil, 0);

  if Count > 0 then
  begin
    DragQueryFile(Msg.Drop, 0, FileName, MAX_PATH);
    Dropped := FileName;

    // URL-Datei (.url)
    if LowerCase(ExtractFileExt(Dropped)) = '.url' then
    begin
      URL := ReadUrlFromInternetShortcut(Dropped);
      if URL <> '' then
      begin
        AddButton(URL, URL);   // Caption = URL, Hint = URL
        SaveButtonsToIni;
      end;
    end

    // LNK → EXE
    else if LowerCase(ExtractFileExt(Dropped)) = '.lnk' then
    begin
      Dropped := ResolveLnk(Dropped);
      AddButton(ChangeFileExt(ExtractFileName(Dropped), ''), Dropped);
      SaveButtonsToIni;
    end

    // EXE oder alles andere...
    else
    begin
      AddButton(ChangeFileExt(ExtractFileName(Dropped), ''), Dropped);
      SaveButtonsToIni;
    end;
  end;

  DragFinish(Msg.Drop);
end;

end.

