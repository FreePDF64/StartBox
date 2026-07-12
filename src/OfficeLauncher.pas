//
// StartBox
//
// Autor: Michael Tesch, Bredstedt
//
// Anfang: 08.06.2026
// Ende:   12.07.2026
//

unit OfficeLauncher;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI,
  Winapi.ShlObj, Vcl.Dialogs, Registry,
  System.SysUtils, System.Classes, System.IniFiles,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.ExtCtrls, Vcl.ImgList, Vcl.Menus, Math,
  System.Generics.Collections, System.ImageList;

type
  TAutorunKind = (akUserRun, akRun);

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
    HotkeyJN: TMenuItem;
    N2: TMenuItem;
    AutostartJN: TMenuItem;
    N3: TMenuItem;
    StartBoxinibearbeiten1: TMenuItem;

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
    procedure AutostartJNClick(Sender: TObject);
    procedure DialogEscHandler(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure StartBoxinibearbeiten1Click(Sender: TObject);
    procedure FormResize(Sender: TObject);

  private
    Ini: TIniFile;
    Buttons: TList<TBitBtn>;
    ForceClose: Boolean;

    OuterMargin: Integer;
    ButtonHeight: Integer;
    ButtonSpacing: Integer;

    IsResizing: Boolean;
    IsDragging: Boolean;
    DragButton: TBitBtn;
    DragOffsetY: Integer;
    DragOffsetX: Integer;
    DragTotalMove: Integer;
    DragStartTime: Cardinal;
    DragDelay: Cardinal;

    OfficeEdition: string;

    PopupButton: TPopupMenu;
    MenuDeleteButton: TMenuItem;
    MenuRenameButton: TMenuItem;
    MenuRunAsAdmin  : TMenuItem;
    MenuEditPath    : TMenuItem;
    ButtonToDelete: TBitBtn;

    procedure WMSysCommand(var Msg: TWMSysCommand); message WM_SYSCOMMAND;
    procedure WMDropFiles(var Msg: TWMDropFiles); message WM_DROPFILES;
    procedure WMHotKey(var Msg: TWMHotKey); message WM_HOTKEY;
    procedure WMExitSizeMove(var Msg: TMessage); message WM_EXITSIZEMOVE;
    procedure CalculateAutoFormSize(out NewW, NewH: Integer);
    procedure ToggleMainForm;

    procedure MinimizeToTray;

    function DetectOfficeEdition: string;
    function FindOfficePath: string;

    function  AddButton(const Caption, ExeName: string): TBitBtn;
    procedure LoadOfficeButtons;
    procedure LoadButtonsFromIni;
    procedure SaveButtonsToIni;
    procedure LoadFormPosition;
    procedure SaveFormPosition;
    function  ButtonExistsForExe(const Exe: string): Boolean;

    procedure RecalculateButtonWidths;
    procedure RecalculateButtonPositions;
    function  CalculateOptimalButtonWidth: Integer;
    function  GetColumnOfButton(Btn: TBitBtn): Integer;

    procedure SwapButtons(A, B: TBitBtn);
    procedure ButtonContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure MenuDeleteButtonClick(Sender: TObject);
    procedure MenuRenameButtonClick(Sender: TObject);
    procedure MenuRunAsAdminClick(Sender: TObject);
    procedure MenuEditPathClick(Sender: TObject);

    function  ResolveLnk(const LnkFile: string): string;

  public
  end;

var
  Form1: TForm1;

var
  DragStartTime: Cardinal;
  DragStartPos: TPoint;
  DragDelay: Cardinal = 120; // ms
  DragThreshold: Integer = 4; // Pixel
  Dragging: Boolean = False;

  Scale: Single;
  ButtonHeight : Integer;
  ButtonSpacing: Integer;
  OuterMargin  : Integer;

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
//  ButtonWidth   = 200;
//  ButtonHeight  = 40;
//  ButtonSpacing = 10;
  HOTKEY_ID     = 1;

// Ist die Hauptform im Vordergrund?
function IsMyWindowForeground(FormHandle: HWND): Boolean;
begin
  Result := GetForegroundWindow = FormHandle;
end;

procedure TForm1.WMExitSizeMove(var Msg: TMessage);
var
  NewW, NewH: Integer;
begin
  inherited;

  if IsResizing then
  begin
    RecalculateButtonPositions;
    CalculateAutoFormSize(NewW, NewH);
    ClientWidth  := NewW;
    ClientHeight := NewH;
    IsResizing := False;
  end;
end;

procedure TForm1.CalculateAutoFormSize(out NewW, NewH: Integer);
var
  I: Integer;
  MaxRight, MaxBottom: Integer;
begin
  MaxRight := 0;
  MaxBottom := 0;

  for I := 0 to Buttons.Count - 1 do
  begin
    MaxRight := Max(MaxRight, Buttons[I].Left + Buttons[I].Width);
    MaxBottom := Max(MaxBottom, Buttons[I].Top + Buttons[I].Height);
  end;

  NewW := MaxRight + OuterMargin;
  NewH := MaxBottom + OuterMargin;
end;

procedure TForm1.ToggleMainForm;
begin
  if IsMyWindowForeground(Form1.Handle) then
  begin
    // Wenn die Form sichtbar und NICHT minimiert ist → minimieren
    if (WindowState <> wsMinimized) and IsWindowVisible(Handle) then
    begin
      MinimizeToTray;
      Exit;
    end;
  end else
  begin
    // Sonst → wiederherstellen + nach vorne holen
    Restore.Click;
    SetForegroundWindow(Handle);      // Fokus erzwingen
    BringWindowToTop(Handle);         // über alle anderen Fenster
  end;
end;

procedure TForm1.DialogEscHandler(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    Key := 0;
    (Sender as TForm).ModalResult := mrNo;
  end;
end;

procedure RunAsAdmin(const FileName, Params: string);
var
  Sei: TShellExecuteInfo;
begin
  ZeroMemory(@Sei, SizeOf(Sei));
  Sei.cbSize := SizeOf(Sei);
  Sei.fMask := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
  Sei.Wnd := 0;
  Sei.lpVerb := 'runas'; // ← erzwingt UAC
  Sei.lpFile := PChar(FileName);
  Sei.lpParameters := PChar(Params);
  Sei.nShow := SW_SHOWNORMAL;

  if not ShellExecuteEx(@Sei) then
    ShowMessage('Starten als Administrator fehlgeschlagen.');
end;

procedure TForm1.WMHotKey(var Msg: TWMHotKey);
begin
  if Msg.HotKey = HOTKEY_ID then
  begin
    ToggleMainForm;
  end;
end;

function TForm1.GetColumnOfButton(Btn: TBitBtn): Integer;
var
  I: Integer;
  ColLeft: Integer;
  MaxWidth: Integer;
begin
  MaxWidth := 0;
  for I := 0 to Buttons.Count - 1 do
    if Buttons[I].Width > MaxWidth then
      MaxWidth := Buttons[I].Width;

  ColLeft := OuterMargin;
  Result := 0;

  while Btn.Left > ColLeft + MaxWidth do
  begin
    Inc(Result);
    ColLeft := ColLeft + MaxWidth + ButtonSpacing;
  end;
end;

procedure TForm1.WMSysCommand(var Msg: TWMSysCommand);
begin
  // Maximieren verhindern
  if (Msg.CmdType = SC_MAXIMIZE) then
    Exit;

  // Minimieren in den Tray
  if (Msg.CmdType and $FFF0) = SC_MINIMIZE then
  begin
    MinimizeToTray;
    Exit;
  end;
  inherited;
end;

// Programm in den Autostart...
function CreateAutorunEntry(const AName, AFilename: string;
  const AKind: TAutorunKind): Boolean;
var
  Reg: TRegistry;
begin
  Result := False;
  Reg := TRegistry.Create;
  try
    if AKind = akUserRun then
      Reg.Rootkey := HKEY_CURRENT_USER
    else
      Reg.Rootkey := HKEY_LOCAL_MACHINE;

    case AKind of
      akRun, akUserRun:
        Result := Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Run', True);
    end;
    Reg.WriteString(AName, AFilename);
  finally
    Reg.Free;
  end;
end;

procedure TForm1.AutostartJNClick(Sender: TObject);
var
  Reg: TRegistry;
begin
  AutostartJN.Checked := Not AutostartJN.Checked;
  if AutostartJN.Checked then
    // Ab in den Autostart nach HKCU: \Software\Microsoft\Windows\CurrentVersion\Run
    CreateAutorunEntry(Application.Title, ParamStr(0), akUserRun)
  else
  // sonst Registry-Eintrag wieder löschen...
  begin
    Reg := TRegistry.Create;
    try
      Reg.Rootkey := HKEY_CURRENT_USER;
      Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Run\', False);
      Reg.DeleteValue('StartBox');
      Reg.CloseKey;
    finally
      Reg.Free;
    end;
  end;
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
    ('StartBox Version 1.5' + #13 + #13 +
     'Copyright © 2026 by FreePDF64@outlook.com' + #13 +
     'Website -> https://github.com/FreePDF64/StartBox' + #13 + #13 +
     'StartBox darf sowohl im privaten als auch im kommerziellen' + #13 +
     'Umfeld ohne Bezahlung eingesetzt werden ("Freeware")!' + #13 +
     'Der Autor übernimmt keinerlei Haftung für Fehler, die direkt' + #13 +
     'oder indirekt aus der Benutzung dieser Software entstehen.' + #13 + #13 +
     'HowTo:' + #13 +
     '- Programme/Dateien/etc. werden per Drag&&Drop hinzugefügt' + #13 +
     '- Icons werden automatisch aus den Dateien extrahiert' + #13 +
     '- Rechtsklick auf Button: u.a. Umbenennen, Löschen, als Admin ausführen' + #13 +
     '- Schließen (X) minimiert in den System Tray' + #13 +
     '- Beenden NUR über das Tray-Menü' + #13 +
     '- Auswahl: Nach Klick auf Buttons minimieren in den System Tray' + #13 +
     '- Auswahl: Maximieren (Toggle) via Hotkey Alt+S' + #13 +
     '- Auswahl: StartBox zum Windows-Autostart hinzufügen',
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

function ExtractShellIconToImageList(const FileName: string; ImageList: TImageList; CurrentPPI: Integer): Integer;
var
  SFI: SHFILEINFO;
  Icon: TIcon;
begin
  Result := -1;

  if SHGetFileInfo(PChar(FileName), 0, SFI, SizeOf(SFI),
     SHGFI_ICON or SHGFI_LARGEICON) = 0 then
    Exit;

  Icon := TIcon.Create;
  try
    Icon.Handle := SFI.hIcon;
    Result := ImageList.AddIcon(Icon);
  finally
    Icon.Free;
  end;

  DestroyIcon(SFI.hIcon);
end;

function TForm1.AddButton(const Caption, ExeName: string): TBitBtn;
var
  Btn: TBitBtn;
  IconIndex: Integer;
  Cap: string;
  Scale: Single;
  W: Integer;
begin
  Scale := CurrentPPI / 96;

  Btn := TBitBtn.Create(Self);
  Btn.Parent := Self;

  // URL?
  if ExeName.StartsWith('http://') or ExeName.StartsWith('https://') then
  begin
    if Caption = '' then
      Cap := ExtractDomainFromURL(ExeName)
    else
      Cap := Caption;

    Btn.Glyph := nil;
    Btn.Images := ImageList1;
    Btn.ImageIndex := 0;
  end
  else
  begin
    Cap := Caption;
    IconIndex := ExtractShellIconToImageList(ExeName, ImageList1, CurrentPPI);
    Btn.Images := ImageList1;
    Btn.ImageIndex := IconIndex;
  end;

  Btn.Caption   := Cap;
  Btn.Hint      := ExeName;
  Btn.ShowHint  := True;
  Btn.Font.Name := 'Segoe UI';

  if Scale <= 1.30 then
    Btn.Font.Size := Max(Round(8.5 * Power(Scale, 0.22)), 8)
  else
    Btn.Font.Size := Max(Round(7 * Power(Scale, 0.22)), 7);

  Btn.Height := ButtonHeight;

  // *** Breite wie alle bestehenden Buttons ***
  W := CalculateOptimalButtonWidth;
  Btn.Width := W;

  Btn.Spacing := ButtonSpacing;
  Btn.Margin  := ButtonSpacing;
  Btn.Layout := blGlyphLeft;

  Btn.PopupMenu := PopupButton;
  Btn.OnContextPopup := ButtonContextPopup;
  Btn.OnMouseDown := ButtonMouseDown;
  Btn.OnMouseMove := ButtonMouseMove;
  Btn.OnMouseUp := ButtonMouseUp;
  Btn.OnClick := ButtonClick;

  // jetzt erst in die Liste
  Buttons.Add(Btn);

  // Positionierung übernimmt deine Spaltenlogik
  RecalculateButtonPositions;

  Result := Btn;
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
  L, T: Integer;
  Btn: TBitBtn;
begin
  Count := Ini.ReadInteger('Buttons', 'Count', 0);

  for I := 0 to Count - 1 do
  begin
    Exe := Ini.ReadString('Buttons', 'Btn' + I.ToString + '_Exe', '');
    Cap := Ini.ReadString('Buttons', 'Btn' + I.ToString + '_Caption', '');

    L := Ini.ReadInteger('Buttons', 'Btn' + I.ToString + '_Left', OuterMargin);
    T := Ini.ReadInteger('Buttons', 'Btn' + I.ToString + '_Top', OuterMargin);

    Btn := AddButton(Cap, Exe);
    Btn.Left := L;
    Btn.Top  := T;
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

    Ini.WriteInteger('Buttons', 'Btn' + I.ToString + '_Left', Buttons[I].Left);
    Ini.WriteInteger('Buttons', 'Btn' + I.ToString + '_Top',  Buttons[I].Top);
  end;
end;

procedure TForm1.LoadFormPosition;
begin
  Left := Ini.ReadInteger('Form', 'Left', Left);
  Top := Ini.ReadInteger('Form', 'Top', Top);

  ClientWidth  := Ini.ReadInteger('Form', 'Width', ClientWidth);
  ClientHeight := Ini.ReadInteger('Form', 'Height', ClientHeight);

  MinimizeJN.Checked := Ini.ReadBool('Form', 'Minimize', MinimizeJN.Checked);
  AutostartJN.Checked := Ini.ReadBool('Autostart', 'Enabled', AutostartJN.Checked);
  HotkeyJN.Checked := Ini.ReadBool('Hotkey', 'Enabled', HotkeyJN.Checked);
end;

procedure TForm1.SaveFormPosition;
begin
  Ini.WriteInteger('Form', 'Left', Left);
  Ini.WriteInteger('Form', 'Top', Top);

  Ini.WriteInteger('Form', 'Width', ClientWidth);
  Ini.WriteInteger('Form', 'Height', ClientHeight);

  Ini.WriteBool('Form', 'Minimize', MinimizeJN.Checked);
  Ini.WriteBool('Autostart', 'Enabled', AutostartJN.Checked);
  Ini.WriteBool('Hotkey', 'Enabled', HotkeyJN.Checked);
end;

procedure TForm1.StartBoxinibearbeiten1Click(Sender: TObject);
var
  StartBoxIni: String;
begin
  StartBoxIni := IncludeTrailingBackslash(ExtractFilePath(Application.ExeName)) + 'StartBox.ini';
  // Wenn die StartBox.ini vorhanden ist, dann...
  if FileExists(StartBoxIni) then
    // Editor aufrufen...
  ShellExecute(Application.Handle, 'open', PChar('notepad.exe'), PChar(' "' + StartBoxIni + '"'), NIL, SW_SHOWNORMAL)
end;

function TForm1.CalculateOptimalButtonWidth: Integer;
var
  I, W: Integer;
begin
  Result := 140; // Mindestbreite

  if Buttons.Count = 0 then
  begin
    Canvas.Font.Assign(Self.Font);
    Exit(Result);
  end;

  Canvas.Font.Assign(Buttons[0].Font);

  for I := 0 to Buttons.Count - 1 do
  begin
    W := Canvas.TextWidth(Buttons[I].Caption) + 90; // 40 Icon + 50 Luft
    if W > Result then
      Result := W;
  end;
end;

procedure TForm1.RecalculateButtonWidths;
var
  I: Integer;
  W, TextW: Integer;
begin
  // Wenn keine Buttons existieren → Default
  if Buttons.Count = 0 then
    Exit;

  // Font für Textmessung
  Canvas.Font.Assign(Buttons[0].Font);

  // Mindestbreite
  W := 120;

  // *** WICHTIG ***
  // Nur bestehende Buttons messen → NICHT den neuen Button
  for I := 0 to Buttons.Count - 1 do
  begin
    TextW := Canvas.TextWidth(Buttons[I].Caption);

    // kompakter Padding: Icon + 12px Luft
    TextW := TextW + 40 + 12;

    if TextW > W then
      W := TextW;
  end;

  // Alle Buttons auf die gleiche Breite setzen
  for I := 0 to Buttons.Count - 1 do
    Buttons[I].Width := W;
end;

procedure TForm1.RecalculateButtonPositions;
var
  I: Integer;
  ColLeft: Integer;
  Y: Integer;
  MaxWidth: Integer;
begin
  if Buttons.Count = 0 then Exit;

  MaxWidth := Buttons[0].Width;

  ColLeft := OuterMargin;
  Y := OuterMargin;

  for I := 0 to Buttons.Count - 1 do
  begin
    // Neue Spalte, wenn unten kein Platz mehr ist
//    if Y + ButtonHeight > ClientHeight - OuterMargin then
    if Y + ButtonHeight > ClientHeight - 7 then
    begin
      ColLeft := ColLeft + MaxWidth + ButtonSpacing;
      Y := OuterMargin;
    end;

    Buttons[I].Left := ColLeft;
    Buttons[I].Top  := Y;

    Y := Y + ButtonHeight + ButtonSpacing;

    if Buttons[I].Width > MaxWidth then
      MaxWidth := Buttons[I].Width;
  end;
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
var
  IniPath: string;
  Sep: TMenuItem;
  Scale: Single;
  IconSize, NewW, NewH: Integer;
  UrlIcon: TIcon;
begin
  Scale := Self.CurrentPPI / 96;

  ButtonHeight  := Round(36 * Scale);
  ButtonSpacing := Round(4 * Scale);
  OuterMargin   := Round(7 * Scale);

  BorderIcons := BorderIcons - [biMaximize];
  CoInitialize(nil);

  Application.ShowHint := True;
  DragDelay := 300;

  IniPath := ChangeFileExt(Application.ExeName, '.ini');
  Ini     := TIniFile.Create(IniPath);
  Buttons := TList<TBitBtn>.Create;

  ForceClose := False;

  if Scale <= 1.30 then
    IconSize := 16
  else
    IconSize := 32;

  UrlIcon := TIcon.Create;
  try
    ImageList1.GetIcon(0, UrlIcon);
  except
    UrlIcon := nil;
  end;

  ImageList1.Clear;
  ImageList1.ColorDepth := cd32bit;
  ImageList1.Masked := False;
  ImageList1.DrawingStyle := dsTransparent;
  ImageList1.BkColor := clNone;
  ImageList1.Width  := IconSize;
  ImageList1.Height := IconSize;

  if Assigned(UrlIcon) then
    ImageList1.AddIcon(UrlIcon);

  UrlIcon.Free;

  PopupButton := TPopupMenu.Create(Self);

  MenuRenameButton := TMenuItem.Create(Self);
  MenuRenameButton.Caption := 'Umbenennen';
  MenuRenameButton.OnClick := MenuRenameButtonClick;
  PopupButton.Items.Add(MenuRenameButton);

  MenuEditPath := TMenuItem.Create(Self);
  MenuEditPath.Caption := 'Pfad/URL anpassen';
  MenuEditPath.OnClick := MenuEditPathClick;
  PopupButton.Items.Add(MenuEditPath);

  MenuDeleteButton := TMenuItem.Create(Self);
  MenuDeleteButton.Caption := 'Löschen';
  MenuDeleteButton.OnClick := MenuDeleteButtonClick;
  PopupButton.Items.Add(MenuDeleteButton);

  Sep := TMenuItem.Create(PopupButton);
  Sep.Caption := '-';
  Sep.Enabled := False;
  PopupButton.Items.Add(Sep);

  MenuRunAsAdmin := TMenuItem.Create(Self);
  MenuRunAsAdmin.Caption := 'Als Administrator ausführen';
  MenuRunAsAdmin.OnClick := MenuRunAsAdminClick;
  PopupButton.Items.Add(MenuRunAsAdmin);

  DragAcceptFiles(Handle, True);

  LoadFormPosition;

  if HotkeyJN.Checked then
    RegisterHotKey(Handle, HOTKEY_ID, MOD_ALT, Ord('S'))
  else
    UnregisterHotKey(Handle, HOTKEY_ID);

  OfficeEdition := DetectOfficeEdition;

  // ---------------------------------------------------------
  // Buttons laden
  // ---------------------------------------------------------
  LoadButtonsFromIni;
  LoadOfficeButtons;

  if Buttons.Count = 0 then
    MessageDlg('Bitte Anwendungen/Dateien/etc. hinzufügen per Drag&Drop.', mtInformation, [mbOK], 0);

  // ---------------------------------------------------------
  // AUTOMATISCHE ANPASSUNG NACH DEM LADEN
  // ---------------------------------------------------------

  // 1) Breite aller Buttons neu berechnen
  RecalculateButtonWidths;

  // 2) Buttons neu anordnen
  RecalculateButtonPositions;

  // 3) Form automatisch anpassen
  CalculateAutoFormSize(NewW, NewH);
  ClientWidth  := NewW;
  ClientHeight := NewH;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  UnregisterHotKey(Handle, HOTKEY_ID);

  SaveFormPosition;
  SaveButtonsToIni;
  Ini.Free;
  Buttons.Free;
  CoUninitialize;
end;

procedure TForm1.FormResize(Sender: TObject);
begin
  if Buttons.Count <= 0 then
    Exit;

  IsResizing := True;
  RecalculateButtonPositions;
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

function HalfOverlap(DragBtn, OtherBtn: TBitBtn): Boolean;
var
  MidY: Integer;
begin
  MidY := DragBtn.Top + (DragBtn.Height div 2);
  Result :=
    (MidY >= OtherBtn.Top) and
    (MidY <= OtherBtn.Top + OtherBtn.Height);
end;

procedure TForm1.SwapButtons(A, B: TBitBtn);
var
  IA, IB: Integer;
  DragMid, TargetMid: Integer;
begin
  IA := Buttons.IndexOf(A);
  IB := Buttons.IndexOf(B);

  DragMid := A.Top + A.Height div 2;
  TargetMid := B.Top + B.Height div 2;

  if DragMid < TargetMid - (B.Height div 4) then
  begin
  end
  else if DragMid > TargetMid + (B.Height div 4) then
  begin
  end
  else
    Exit;

  Buttons[IA] := B;
  Buttons[IB] := A;

  RecalculateButtonPositions;
end;

procedure TForm1.ButtonMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button <> mbLeft then Exit;

  DragButton := TBitBtn(Sender);

  DragStartTime := GetTickCount;
  DragStartPos := Point(X, Y);

  DragOffsetX := X;
  DragOffsetY := Y;

  DragTotalMove := 0;

  IsDragging := False; // wird in MouseMove aktiviert
end;

procedure TForm1.ButtonMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var
  Btn: TBitBtn;
  I: Integer;
  Other: TBitBtn;
  Col: Integer;
  P: TPoint;
begin
  Btn := DragButton;
  if Btn = nil then Exit;

  // Drag aktivieren
  if not IsDragging then
  begin
    if (GetTickCount - DragStartTime >= DragDelay) and
       (Abs(X - DragStartPos.X) + Abs(Y - DragStartPos.Y) > 3) then
      IsDragging := True
    else
      Exit;
  end;

  // Maus absolut holen
  GetCursorPos(P);
  P := ScreenToClient(P);

  // Button bewegen (zitternfrei)
  Btn.Left := P.X - DragOffsetX;
  Btn.Top  := P.Y - DragOffsetY;

  // Spalte bestimmen
  Col := GetColumnOfButton(Btn);

  // Buttons in derselben Spalte prüfen
  for I := 0 to Buttons.Count - 1 do
  begin
    Other := Buttons[I];
    if Other = Btn then Continue;

    // Nur Buttons in derselben Spalte
    if GetColumnOfButton(Other) <> Col then Continue;

    // Überlappung → originaler schöner Swap
    if Abs((Btn.Top + Btn.Height div 2) - (Other.Top + Other.Height div 2)) < (Btn.Height div 2) then
    begin
      SwapButtons(Btn, Other);
      Break; // nur einen Button tauschen
    end;
  end;
end;

procedure TForm1.ButtonMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if (Button = mbLeft) and IsDragging then
    SaveButtonsToIni;

  IsDragging := False;
  DragButton := nil;

  RecalculateButtonPositions;
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

procedure TForm1.MenuRunAsAdminClick(Sender: TObject);
var
  dlg: TForm;
  lbl: TLabel;
  btnYes, btnNo: TButton;
  p: TPoint;
  res: Integer;
  Scale: Single;
  MsgText: string;
  ExePath: string;
  BtnW, BtnH, Spacing: Integer;
  PaddingY: Integer;
  TextH: Integer;
  BottomPadding: Integer;
  LabelBottom: Integer;
begin
  if ButtonToDelete = nil then Exit;

  // EXE-Pfad aus dem Button holen
  ExePath := ButtonToDelete.Hint;
  if ExePath = '' then Exit;

  // DPI-Faktor
  Scale := Self.CurrentPPI / 96;

  // Basiswerte
  BtnW := Max(Round(90 * Scale), 90);
  Spacing := Round(10 * Scale);
  PaddingY := Round(4 * Scale);
  BottomPadding := Round(10 * Scale);

  // Text dynamisch erzeugen
  MsgText := 'Als Administrator ausführen:' + '   ' + sLineBreak + '"' + ButtonToDelete.Caption + '"';

  dlg := TForm.Create(Self);
  try
    dlg.KeyPreview := True;
    dlg.OnKeyDown := DialogEscHandler;

    dlg.BorderStyle := bsDialog;
    dlg.Caption := 'Administrator';
    dlg.Position := poDesigned;

    dlg.ClientWidth := Max(Round(280 * Scale), 280);

    // Label
    lbl := TLabel.Create(dlg);
    lbl.Parent := dlg;
    lbl.Caption := MsgText;
    lbl.Left := Round(10 * Scale);
    lbl.Top  := Round(10 * Scale);
    lbl.Width := dlg.ClientWidth - Round(20 * Scale);
    lbl.WordWrap := True;

    // Ja-Button
    btnYes := TButton.Create(dlg);
    btnYes.Parent := dlg;
    btnYes.Caption := 'Ja';
    btnYes.ModalResult := mrYes;
    btnYes.Width := BtnW;

    // Nein-Button
    btnNo := TButton.Create(dlg);
    btnNo.Parent := dlg;
    btnNo.Caption := 'Nein';
    btnNo.ModalResult := mrNo;
    btnNo.Width := BtnW;

    // Texthöhe über Font.Height
    TextH := Abs(btnYes.Font.Height);

    // Automatische Buttonhöhe
    BtnH := Max(TextH + PaddingY * 2, 24);

    btnYes.Height := BtnH;
    btnNo.Height := BtnH;

    btnYes.Padding.Top := PaddingY;
    btnYes.Padding.Bottom := PaddingY;

    btnNo.Padding.Top := PaddingY;
    btnNo.Padding.Bottom := PaddingY;

    // Unterkante des Labels bestimmen
    LabelBottom := lbl.Top + lbl.Height;

    // Buttons direkt unter dem Label, rechtsbündig
    btnNo.Left := dlg.ClientWidth - BtnW - Spacing;
    btnNo.Top  := LabelBottom + (Spacing);  // Abstand vergrößert

    btnYes.Left := btnNo.Left - BtnW - Spacing;
    btnYes.Top  := btnNo.Top;

    // Dialoghöhe exakt anpassen → unterer Rand kurz unter Buttons
    dlg.ClientHeight := btnNo.Top + BtnH + BottomPadding;

    // Dialog unterhalb des Buttons anzeigen
    p := ButtonToDelete.ClientToScreen(Point(0, 0));
    dlg.Left := p.X;
    dlg.Top := p.Y + ButtonToDelete.Height + Round(10 * Scale);

    KeepFormOnScreen(dlg);
    res := dlg.ShowModal;

    if res <> mrYes then Exit;

    // Jetzt als Administrator starten
    RunAsAdmin(ExePath, '');

  finally
    dlg.Free;
  end;
end;

procedure TForm1.MenuDeleteButtonClick(Sender: TObject);
var
  dlg: TForm;
  lbl: TLabel;
  btnYes, btnNo: TButton;
  p: TPoint;
  res: Integer;
  Scale: Single;
  MsgText: string;
  BtnW, BtnH, Spacing: Integer;
  PaddingY: Integer;
  TextH: Integer;
  BottomPadding: Integer;
  LabelBottom: Integer;
  NewW, NewH: Integer;   // <‑‑ hinzugefügt
begin
  if ButtonToDelete = nil then Exit;

  // DPI-Faktor
  Scale := Self.CurrentPPI / 96;

  // Basiswerte
  BtnW := Max(Round(90 * Scale), 90);
  Spacing := Round(10 * Scale);
  PaddingY := Round(4 * Scale);
  BottomPadding := Round(10 * Scale);

  // Text dynamisch erzeugen
  MsgText := 'Button löschen:' + '   ' + sLineBreak + '"' + ButtonToDelete.Caption + '"';

  dlg := TForm.Create(Self);
  try
    dlg.KeyPreview := True;
    dlg.OnKeyDown := DialogEscHandler;

    dlg.BorderStyle := bsDialog;
    dlg.Caption := 'Löschen';
    dlg.Position := poDesigned;

    dlg.ClientWidth := Max(Round(280 * Scale), 280);

    // Label
    lbl := TLabel.Create(dlg);
    lbl.Parent := dlg;
    lbl.Caption := MsgText;
    lbl.Left := Round(10 * Scale);
    lbl.Top  := Round(10 * Scale);
    lbl.Width := dlg.ClientWidth - Round(20 * Scale);
    lbl.WordWrap := True;

    // Ja-Button
    btnYes := TButton.Create(dlg);
    btnYes.Parent := dlg;
    btnYes.Caption := 'Ja';
    btnYes.ModalResult := mrYes;
    btnYes.Width := BtnW;

    // Nein-Button
    btnNo := TButton.Create(dlg);
    btnNo.Parent := dlg;
    btnNo.Caption := 'Nein';
    btnNo.ModalResult := mrNo;
    btnNo.Width := BtnW;

    // Texthöhe über Font.Height
    TextH := Abs(btnYes.Font.Height);

    // Automatische Buttonhöhe
    BtnH := Max(TextH + PaddingY * 2, 24);

    btnYes.Height := BtnH;
    btnNo.Height := BtnH;

    btnYes.Padding.Top := PaddingY;
    btnYes.Padding.Bottom := PaddingY;

    btnNo.Padding.Top := PaddingY;
    btnNo.Padding.Bottom := PaddingY;

    // Unterkante des Labels bestimmen
    LabelBottom := lbl.Top + lbl.Height;

    // Buttons direkt unter dem Label, rechtsbündig
    btnNo.Left := dlg.ClientWidth - BtnW - Spacing;
    btnNo.Top  := LabelBottom + (Spacing);

    btnYes.Left := btnNo.Left - BtnW - Spacing;
    btnYes.Top  := btnNo.Top;

    // Dialoghöhe exakt anpassen
    dlg.ClientHeight := btnNo.Top + BtnH + BottomPadding;

    // Dialog unterhalb des Buttons anzeigen
    p := ButtonToDelete.ClientToScreen(Point(0, 0));
    dlg.Left := p.X;
    dlg.Top := p.Y + ButtonToDelete.Height + Round(10 * Scale);

    KeepFormOnScreen(dlg);
    res := dlg.ShowModal;

    if res <> mrYes then Exit;

    // ---------------------------------------------------------
    // BUTTON LÖSCHEN + FORM NEU BERECHNEN
    // ---------------------------------------------------------

    Buttons.Remove(ButtonToDelete);
    ButtonToDelete.Free;
    ButtonToDelete := nil;

    // 1) Breite aller Buttons neu berechnen
    RecalculateButtonWidths;

    // 2) Buttons neu anordnen
    RecalculateButtonPositions;

    // 3) Form automatisch anpassen
    CalculateAutoFormSize(NewW, NewH);
    ClientWidth  := NewW;
    ClientHeight := NewH;

    // 4) Speichern
    SaveButtonsToIni;

  finally
    dlg.Free;
  end;
end;

procedure TForm1.MenuEditPathClick(Sender: TObject);
var
  dlg: TForm;
  edt: TEdit;
  lbl: TLabel;
  btnOK, btnCancel: TButton;
  p: TPoint;
  res: Integer;
  Scale: Single;
  BtnW, BtnH, Spacing: Integer;
  PaddingY: Integer;
  TextH: Integer;
  BottomPadding: Integer;
begin
  if ButtonToDelete = nil then Exit;

  Scale := Self.CurrentPPI / 96;

  BtnW := Max(Round(90 * Scale), 90);
  Spacing := Round(10 * Scale);
  PaddingY := Round(4 * Scale);
  BottomPadding := Round(10 * Scale);

  dlg := TForm.Create(Self);
  try
    dlg.KeyPreview := True;
    dlg.OnKeyDown := DialogEscHandler;

    dlg.BorderStyle := bsDialog;
    dlg.Caption := 'Pfad/URL anpassen';
    dlg.Position := poDesigned;

    dlg.ClientWidth := Max(Round(280 * Scale), 280);

    // Label
    lbl := TLabel.Create(dlg);
    lbl.Parent := dlg;
    lbl.Caption := 'Neuer Pfad / neue URL:';
    lbl.Left := Round(10 * Scale);
    lbl.Top  := Round(10 * Scale);

    // Eingabefeld
    edt := TEdit.Create(dlg);
    edt.Parent := dlg;
    edt.Left := Round(10 * Scale);
    edt.Top  := Round(35 * Scale);
    edt.Width := dlg.ClientWidth - Round(20 * Scale);
    edt.Height := Max(Round(26 * Scale), 26);
    edt.Text := ButtonToDelete.Hint;   // <‑‑ WICHTIG

    // OK
    btnOK := TButton.Create(dlg);
    btnOK.Parent := dlg;
    btnOK.Caption := 'OK';
    btnOK.ModalResult := mrOk;
    btnOK.Width := BtnW;

    // Cancel
    btnCancel := TButton.Create(dlg);
    btnCancel.Parent := dlg;
    btnCancel.Caption := 'Abbrechen';
    btnCancel.ModalResult := mrCancel;
    btnCancel.Width := BtnW;

    TextH := Abs(btnOK.Font.Height);
    BtnH := Max(TextH + PaddingY * 2, 24);

    btnOK.Height := BtnH;
    btnCancel.Height := BtnH;

    btnOK.Padding.Top := PaddingY;
    btnOK.Padding.Bottom := PaddingY;
    btnCancel.Padding.Top := PaddingY;
    btnCancel.Padding.Bottom := PaddingY;

    // Buttons unter dem Edit
    btnCancel.Left := dlg.ClientWidth - BtnW - Spacing;
    btnCancel.Top  := edt.Top + edt.Height + Spacing;

    btnOK.Left := btnCancel.Left - BtnW - Spacing;
    btnOK.Top  := btnCancel.Top;

    dlg.ClientHeight := btnCancel.Top + BtnH + BottomPadding;

    // Dialog unterhalb des Buttons anzeigen
    p := ButtonToDelete.ClientToScreen(Point(0, 0));
    dlg.Left := p.X;
    dlg.Top := p.Y + ButtonToDelete.Height + Round(10 * Scale);

    KeepFormOnScreen(dlg);
    res := dlg.ShowModal;

    if res <> mrOk then Exit;
    if Trim(edt.Text) = '' then Exit;

    // Pfad/URL aktualisieren
    ButtonToDelete.Hint := edt.Text;

    // Buttons neu berechnen
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
  Scale: Single;
  BtnW, BtnH, Spacing: Integer;
  PaddingY: Integer;
  TextH: Integer;
  BottomPadding: Integer;
begin
  if ButtonToDelete = nil then Exit;

  // DPI-Faktor
  Scale := Self.CurrentPPI / 96;

  // Basiswerte
  BtnW := Max(Round(90 * Scale), 90);
  Spacing := Round(10 * Scale);
  PaddingY := Round(4 * Scale);
  BottomPadding := Round(10 * Scale);  // kleiner Abstand zum unteren Rand

  dlg := TForm.Create(Self);
  try
    dlg.KeyPreview := True;
    dlg.OnKeyDown := DialogEscHandler;

    dlg.BorderStyle := bsDialog;
    dlg.Caption := 'Umbenennen';
    dlg.Position := poDesigned;

    dlg.ClientWidth := Max(Round(280 * Scale), 280);

    // Label
    lbl := TLabel.Create(dlg);
    lbl.Parent := dlg;
    lbl.Caption := 'Neuer Name:';
    lbl.Left := Round(10 * Scale);
    lbl.Top  := Round(10 * Scale);

    // Eingabefeld
    edt := TEdit.Create(dlg);
    edt.Parent := dlg;
    edt.Left := Round(10 * Scale);
    edt.Top  := Round(35 * Scale);
    edt.Width := dlg.ClientWidth - Round(20 * Scale);
    edt.Height := Max(Round(26 * Scale), 26);
    edt.Text := ButtonToDelete.Caption;

    // OK-Button
    btnOK := TButton.Create(dlg);
    btnOK.Parent := dlg;
    btnOK.Caption := 'OK';
    btnOK.ModalResult := mrOk;
    btnOK.Width := BtnW;

    // Cancel-Button
    btnCancel := TButton.Create(dlg);
    btnCancel.Parent := dlg;
    btnCancel.Caption := 'Abbrechen';
    btnCancel.ModalResult := mrCancel;
    btnCancel.Width := BtnW;

    // Texthöhe über Font.Height
    TextH := Abs(btnOK.Font.Height);

    // Automatische Buttonhöhe
    BtnH := Max(TextH + PaddingY * 2, 24);

    btnOK.Height := BtnH;
    btnCancel.Height := BtnH;

    btnOK.Padding.Top := PaddingY;
    btnOK.Padding.Bottom := PaddingY;

    btnCancel.Padding.Top := PaddingY;
    btnCancel.Padding.Bottom := PaddingY;

    // Buttons direkt unter TEdit, rechtsbündig
    btnCancel.Left := dlg.ClientWidth - BtnW - Spacing;
    btnCancel.Top  := edt.Top + edt.Height + Spacing;

    btnOK.Left := btnCancel.Left - BtnW - Spacing;
    btnOK.Top  := btnCancel.Top;

    // Dialoghöhe exakt anpassen → unterer Rand kurz unter Buttons
    dlg.ClientHeight :=
      btnCancel.Top + BtnH + BottomPadding;

    // Dialog unterhalb des Buttons anzeigen
    p := ButtonToDelete.ClientToScreen(Point(0, 0));
    dlg.Left := p.X;
    dlg.Top := p.Y + ButtonToDelete.Height + Round(10 * Scale);

    KeepFormOnScreen(dlg);
    res := dlg.ShowModal;

    if res <> mrOk then Exit;
    if Trim(edt.Text) = '' then Exit;

    ButtonToDelete.Caption := edt.Text;

    RecalculateButtonWidths;
    RecalculateButtonPositions;

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
  Dropped, URL, URL2: string;
  NewW, NewH: Integer;
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
        URL2 := ExtractDomainFromURL(URL);
        AddButton(URL2, URL);   // Caption = URL2, Hint = URL
      end;
    end

    // LNK → EXE
    else if LowerCase(ExtractFileExt(Dropped)) = '.lnk' then
    begin
      Dropped := ResolveLnk(Dropped);
      AddButton(ChangeFileExt(ExtractFileName(Dropped), ''), Dropped);
    end

    // EXE oder alles andere...
    else
    begin
      AddButton(ChangeFileExt(ExtractFileName(Dropped), ''), Dropped);
    end;

    // *** WICHTIG ***
    // Nach jedem Hinzufügen:
    // 1) Breite aller Buttons neu berechnen
    RecalculateButtonWidths;

    // 2) Buttons neu anordnen
    RecalculateButtonPositions;

    // 3) Form automatisch anpassen
    CalculateAutoFormSize(NewW, NewH);
    ClientWidth  := NewW;
    ClientHeight := NewH;

    // 4) Speichern
    SaveButtonsToIni;
  end;

  DragFinish(Msg.Drop);
end;


end.

