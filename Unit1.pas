unit Unit1;

interface

uses
    pipe,
    Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
    System.Classes, Vcl.Graphics, System.DateUtils,
    Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.ExtCtrls,
    Vcl.ComCtrls, System.Generics.Collections, Vcl.Menus, VclTee.TeeGDIPlus,
    VclTee.TeEngine, VclTee.TeeProcs, VclTee.Chart, VclTee.Series,
    System.ImageList, Vcl.ImgList;

type
    TReadPipeThread = class(TThread)
    private
        procedure Execute; override;

    end;

    TParty = Record
        id: int64;
        CreatedAt: TDateTime;

    end;

    TProduct = Record
        id: int64;
        serial, order: integer;

        function What: string;
    end;

    TForm1 = class(TForm)
        PageControl1: TPageControl;
        TabSheet1: TTabSheet;
        TabSheet2: TTabSheet;
        Panel4: TPanel;
        Panel3: TPanel;
        TreeView1: TTreeView;
        Panel1: TPanel;
        Chart1: TChart;
        Panel5: TPanel;
        Chart2: TChart;
        Panel7: TPanel;
        Panel8: TPanel;
        Panel9: TPanel;
        PanelBottomMessage: TPanel;
        Panel11: TPanel;
        ListView2: TListView;
        ImageList1: TImageList;
        Panel10: TPanel;
        LabelConnectComPort: TLabel;
        Button1: TButton;
        Button2: TButton;
        ComboBox1: TComboBox;
        procedure FormCreate(Sender: TObject);
        procedure ListView1CustomDrawSubItem(Sender: TCustomListView;
          Item: TListItem; SubItem: integer; State: TCustomDrawState;
          var DefaultDraw: Boolean);
        procedure ListView1ItemChecked(Sender: TObject; Item: TListItem);
        procedure ListView1SelectItem(Sender: TObject; Item: TListItem;
          Selected: Boolean);
        procedure TreeView1Expanding(Sender: TObject; Node: TTreeNode;
          var AllowExpansion: Boolean);
        procedure TreeView1Change(Sender: TObject; Node: TTreeNode);
        procedure ListView2Edited(Sender: TObject; Item: TListItem;
          var s: string);
        procedure ListView2MouseDown(Sender: TObject; Button: TMouseButton;
          Shift: TShiftState; X, Y: integer);
        procedure Button2Click(Sender: TObject);
        procedure ComboBox1Change(Sender: TObject);
        procedure Button1Click(Sender: TObject);
        procedure ListView2CustomDrawItem(Sender: TCustomListView;
          Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
        procedure FormActivate(Sender: TObject);
        procedure FormClose(Sender: TObject; var Action: TCloseAction);
    private
        { Private declarations }
        treenode_party: TDictionary<TTreeNode, TParty>;
        treenode_product: TDictionary<TTreeNode, TProduct>;
        current_products: TList<TProduct>;
        place_error: array [0 .. 9] of Boolean;

        procedure OnException(Sender: TObject; E: Exception);

        procedure Handle_data_from_master(cmd: integer);

        procedure Handle_years;
        procedure Handle_months_of_year;
        procedure Handle_days_of_year_month;
        procedure Handle_parties_of_year_month_day;
        procedure Handle_products_of_party;
        procedure Handle_sensitivities_of_product;
        procedure Handle_current_party;
        procedure Handle_hardware_config;
        procedure Handle_hardware_sensitivity;

        function NodeYear(year: integer): TTreeNode;
        function NodeYearMonth(year, month: integer): TTreeNode;
        function NodeYearMonthDay(year, month, day: integer): TTreeNode;
        function NodeParty(partyID: int64): TTreeNode;
        function NodeProduct(productID: int64): TTreeNode;

        // procedure WMREGCHANGE(var Msg: TMessage); message WM_REGCHANGE;

    public
        { Public declarations }

    end;

var
    Form1: TForm1;
    pipe_write: TPipeClient;
    pipe_read: TPipeClient;
    read_pipe_thread: TReadPipeThread;

implementation

uses System.IOUtils;

const
    ACTION_YEARS = 0;
    ACTION_MONTHS_OF_YEAR = 1;
    ACTION_DAYS_OF_YEAR_MONTH = 2;
    ACTION_PARTIES_OF_YEAR_MONTH_DAY = 3;
    ACTION_PRODUCTS_OF_PARTY = 4;
    ACTION_SENSITIVITIES_OF_PRODUCT = 5;
    ACTION_CURRENT_PARTY = 6;
    ACTION_INFO_MESSAGE = 7;
    ACTION_HARDWARE_SENSITIVITY = 8;
    ACTION_HARDWARE_CONNECTED = 9;
    ACTION_HARDWARE_DISCONNECTED = 10;
    ACTION_HARDWARE_CONNECTION_ERROR = 11;
    ACTION_HARDWARE_CONFIG = 12;
    ACTION_HARDWARE_CURRENT_PLACE = 13;
    ACTION_COM_PORTS = 14;

    MSG_YEARS = 0;
    MSG_MONTHS_OF_YEAR = 1;
    MSG_DAYS_OF_YEAR_MONTH = 2;
    MSG_PARTIES_OF_YEAR_MONTH_DAY = 3;
    MSG_PRODUCTS_OF_PARTY = 4;
    MSG_SENSITIVITIES_OF_PRODUCT = 5;
    MSG_CURRENT_PRODUCT_SERIAL = 6;
    MSG_CREATE_NEW_PARTY = 7;

    MSG_PORT_NAME = 8;
    MSG_PLACE_CHECKED = 9;
    MSG_START_HARDWARE = 10;
    MSG_STOP_HARDWARE = 11;

procedure TReadPipeThread.Execute;
begin
    while not Terminated do
    begin
        Form1.Handle_data_from_master(pipe_read.ReadInt32);
    end;
end;

procedure terminate_error(error_text: string);
var
    f: TextFile;
    col, row: integer;
begin
    AssignFile(f, ExtractFileDir(paramstr(0)) + '\fail.txt');
    ReWrite(f);
    WriteLn(f, error_text);
    CloseFile(f);
    ExitProcess(1);
end;

function month_number_to_str(n: integer): string;
begin
    result := FormatDateTime('mmmm', EncodeDateTime(2000, n, 1, 0, 0, 0, 0));
end;

function day_to_str(n: integer): string;
begin
    result := inttostr(n);
    if n < 10 then
        result := '0' + result;
end;

function read_party: TParty;
begin
    result.id := pipe_read.ReadInt64;
    result.CreatedAt := pipe_read.ReadDateTime;
end;

function read_product: TProduct;
begin
    result.id := pipe_read.ReadInt64;
    result.order := pipe_read.ReadInt32;
    result.serial := pipe_read.ReadInt32;
end;

function read_products: TList<TProduct>;
var
    i: integer;
begin
    result := TList<TProduct>.create;
    for i := 0 to pipe_read.ReadInt32 - 1 do
    begin
        result.Add(read_product);
    end;
end;

procedure combobox_comports_update(combobox: TComboBox; ports: TStringList);
var
    s: string;
begin
    s := combobox.Text;
    combobox.Items.Assign(ports);
    combobox.ItemIndex := combobox.Items.IndexOf(s);
    if (combobox.Items.IndexOf(s) = -1) and (ports.Count > 0) then
    begin
        combobox.ItemIndex := 0;
    end;
end;

function TProduct.What;
begin
    result := Format('%.*d: %d', [2, order + 1, serial])
end;

procedure TForm1.Handle_data_from_master(cmd: integer);
var
    i, n, n1, o: integer;
    Y: double;
    s, s1: string;
    xs: array [0 .. 9] of integer;
    strs: TStringList;
begin
    case cmd of
        ACTION_YEARS:
            Handle_years;
        ACTION_MONTHS_OF_YEAR:
            Handle_months_of_year;
        ACTION_DAYS_OF_YEAR_MONTH:
            Handle_days_of_year_month;
        ACTION_PARTIES_OF_YEAR_MONTH_DAY:
            Handle_parties_of_year_month_day;
        ACTION_PRODUCTS_OF_PARTY:
            Handle_products_of_party;
        ACTION_SENSITIVITIES_OF_PRODUCT:
            Handle_sensitivities_of_product;
        ACTION_CURRENT_PARTY:
            Handle_current_party;
        ACTION_INFO_MESSAGE:
            begin
                s := pipe_read.ReadString;
                s1 := pipe_read.ReadString;
                read_pipe_thread.Synchronize(
                    procedure
                    begin
                        PanelBottomMessage.Caption := s;
                        PanelBottomMessage.Font.Color := StringToColor(s1);
                    end);
            end;
        ACTION_HARDWARE_SENSITIVITY:
            Handle_hardware_sensitivity;

        ACTION_HARDWARE_CONNECTED:
            begin
                read_pipe_thread.Synchronize(
                    procedure
                    var
                        i: integer;
                    begin
                        Button1.Caption := 'Осстановить опрос';
                        for i := 0 to 9 do
                        begin
                            Chart2.Series[i].Clear;
                        end;
                        LabelConnectComPort.Caption := 'cвязь установлена';
                        LabelConnectComPort.Font.Color := clNavy;
                    end);
            end;

        ACTION_HARDWARE_DISCONNECTED:
            begin
                read_pipe_thread.Synchronize(
                    procedure
                    begin
                        Button1.Caption := 'Запустить опрос';
                        LabelConnectComPort.Caption := 'опрос прерван';
                        LabelConnectComPort.Font.Color := clMaroon;
                    end);
            end;

        ACTION_HARDWARE_CONFIG:
            Handle_hardware_config;
        ACTION_HARDWARE_CONNECTION_ERROR:
            begin
                s := pipe_read.ReadString;
                read_pipe_thread.Synchronize(
                    procedure
                    begin
                        LabelConnectComPort.Caption := s;
                        LabelConnectComPort.Font.Color := clREd;

                    end);
            end;
        ACTION_HARDWARE_CURRENT_PLACE:
            begin
                n := pipe_read.ReadInt32;
                read_pipe_thread.Synchronize(
                    procedure
                    var
                        i: integer;
                    begin
                        for i := 0 to 9 do
                        begin
                            if ListView2.Items[i].ImageIndex = 1 then
                                ListView2.Items[i].ImageIndex := -1;
                        end;
                        if (n > -1) and (n < 10) then
                            ListView2.Items[n].ImageIndex := 1;
                    end);
            end;
        ACTION_COM_PORTS:
            begin
                strs := TStringList.create;
                for i := 0 to pipe_read.ReadInt32 - 1 do
                begin
                    strs.Add(pipe_read.ReadString);
                end;
                read_pipe_thread.Synchronize(
                    procedure
                    begin
                        combobox_comports_update(ComboBox1, strs);
                    end);

                strs.Free;
            end;

    end;

end;

procedure TForm1.Handle_hardware_sensitivity;
var
    n, status, code: integer;
    Y: double;
    errStr, s: string;
begin
    n := pipe_read.ReadInt32;
    status := pipe_read.ReadInt32;
    Y := pipe_read.ReadFloat64;
    errStr := pipe_read.ReadString;

    read_pipe_thread.Synchronize(
        procedure
        var
            s: string;
            xs: TStrings;
        begin
            place_error[n] := errStr <> '';
            if (errStr = '') and (status = 0) then
            begin
                s := FormatFloat('#0.###', Y);
                ListView2.Items[n].SubItems[1] := s;
                Chart2.Series[n].AddXY(Now(), Y);
                PanelBottomMessage.Font.Color := clNavy;
                PanelBottomMessage.Caption := Format('%d: %s', [n + 1, s]);
                ListView2.Items[n].ImageIndex := -1;
            end;
            if (errStr <> '') then
            begin
                ListView2.Items[n].SubItems[1] := errStr;
                ListView2.Items[n].ImageIndex := 0;
                PanelBottomMessage.Font.Color := clREd;
                PanelBottomMessage.Caption := Format('%d: %s', [n + 1, errStr]);

            end;
            s := ListView2.Items[n].Caption;
            ListView2.Items[n].Caption := '';
            ListView2.Items[n].Caption := s;
            xs := ListView2.Items[n].SubItems;
            xs[0] := xs[0];
            xs[1] := xs[1];
            xs[2] := xs[2];

        end)
end;

procedure TForm1.Handle_hardware_config;
var
    i: integer;
    s: string;
    xs: array [0 .. 9] of integer;

begin

    s := pipe_read.ReadString;
    for i := 0 to 9 do
    begin
        xs[i] := pipe_read.ReadInt32;
    end;

    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
        begin
            for i := 0 to ComboBox1.Items.Count - 1 do
            begin
                if ComboBox1.Items[i] = s then
                begin
                    ComboBox1.OnChange := nil;
                    ComboBox1.ItemIndex := i;
                    ComboBox1.OnChange := ComboBox1Change;
                    Break;
                end;
            end;
            for i := 0 to 9 do
            begin
                ListView2.Items[i].Checked := xs[i] <> 0;
                Chart2.Series[i].Active := ListView2.Items[i].Checked;
            end;
        end);

end;

procedure TForm1.Handle_years;
var
    i: integer;
    years: TList<integer>;
begin
    years := TList<integer>.create;
    for i := 0 to pipe_read.ReadInt32 - 1 do
    begin
        years.Add(pipe_read.ReadInt32);
    end;
    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
            Node: TTreeNode;
        begin
            TreeView1.Items.Clear;
            treenode_party.Clear;
            treenode_product.Clear;
            for i := 0 to years.Count - 1 do
            begin
                Node := TreeView1.Items.Add(nil, inttostr(years[i]));
                TreeView1.Items.AddChild(Node, '');
            end;
        end);
    years.Free;
end;

procedure TForm1.Handle_months_of_year;
var
    year, n, i: integer;
    months: TList<integer>;
begin
    year := pipe_read.ReadInt32;
    months := TList<integer>.create;
    for i := 0 to pipe_read.ReadInt32 - 1 do
    begin
        months.Add(pipe_read.ReadInt32);
    end;

    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
            Node, node_year: TTreeNode;
        begin
            node_year := NodeYear(year);
            TreeView1.Items.Delete(node_year.getFirstChild);
            for i := 0 to months.Count - 1 do
            begin
                Node := TreeView1.Items.AddChild(node_year,
                  month_number_to_str(months[i]));
                TreeView1.Items.AddChild(Node, '');
            end;
        end);

    months.Free;
end;

procedure TForm1.Handle_days_of_year_month;
var
    year, month, n, i: integer;
    days: TList<integer>;
begin
    year := pipe_read.ReadInt32;
    month := pipe_read.ReadInt32;
    days := TList<integer>.create;
    for i := 0 to pipe_read.ReadInt32 - 1 do
    begin
        days.Add(pipe_read.ReadInt32);
    end;

    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
            Node, node_year_month: TTreeNode;

        begin
            node_year_month := NodeYearMonth(year, month);
            TreeView1.Items.Delete(node_year_month.getFirstChild);
            for i := 0 to days.Count - 1 do
            begin
                Node := TreeView1.Items.AddChild(node_year_month,
                  day_to_str( days[i]));
                TreeView1.Items.AddChild(Node, '');
            end;
        end);

    days.Free;
end;

procedure TForm1.Handle_parties_of_year_month_day;
var
    year, month, day, n, i, Count: integer;
    parties: TList<TParty>;
begin
    year := pipe_read.ReadInt32;
    month := pipe_read.ReadInt32;
    day := pipe_read.ReadInt32;
    parties := TList<TParty>.create;
    Count := pipe_read.ReadInt32;
    for i := 0 to Count - 1 do
    begin
        parties.Add(read_party);
    end;

    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
            Node_YearMonthDay, Node: TTreeNode;
            p: TParty;
        begin
            Node_YearMonthDay := NodeYearMonthDay(year, month, day);
            TreeView1.Items.Delete(Node_YearMonthDay.getFirstChild);
            for i := 0 to parties.Count - 1 do
            begin
                p := parties[i];
                Node := TreeView1.Items.AddChild(Node_YearMonthDay,
                  FormatDateTime('hh:nn:ss', p.CreatedAt));
                treenode_party.Add(Node, p);
                TreeView1.Items.AddChild(Node, '');

            end;
        end);
    parties.Free;
end;

procedure TForm1.Handle_products_of_party;
var
    i, Count, order, serial: integer;
    party: TParty;
    products: TList<TProduct>;
begin
    party := read_party;
    products := read_products;
    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
            Node_party, Node: TTreeNode;
            p: TProduct;
        begin
            Node_party := NodeParty(party.id);
            TreeView1.Items.Delete(Node_party.getFirstChild);
            for i := 0 to products.Count - 1 do
            begin
                p := products[i];
                Node := TreeView1.Items.AddChild(Node_party, p.What);
                treenode_product.Add(Node, p);
            end;

        end);
    products.Free;

end;

procedure TForm1.Handle_sensitivities_of_product;
var
    n_product, i, n_series, product_Serial: integer;
    product_id: int64;
    xs: TList<TDateTime>;
    ys: TList<double>;
begin

    xs := TList<TDateTime>.create();
    ys := TList<double>.create();
    product_id := pipe_read.ReadInt64;
    for i := 0 to pipe_read.ReadInt32 - 1 do
    begin
        xs.Add(pipe_read.ReadDateTime);
        ys.Add(pipe_read.ReadFloat64);
    end;
    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
        begin
            Chart1.Hide;
            Chart1.Series[0].Clear;
            if xs.Count > 0 then
            begin
                Chart1.Title.Caption :=
                  FormatDateTime('dddd dd mmmm yyyy hh:mm:ss', xs[0]);
            end;

            for i := 0 to xs.Count - 1 do
            begin
                Chart1.Series[0].Active := true;
                Chart1.Series[0].AddXY(xs[i], ys[i]);
            end;
            Chart1.Show;

        end);

    xs.Free;
    ys.Free;
end;

procedure TForm1.Handle_current_party;
var
    party: TParty;
    i: integer;
begin
    if current_products <> nil then
    begin
        current_products.Free;
    end;

    party := read_party;
    current_products := read_products;

    read_pipe_thread.Synchronize(
        procedure
        var
            i: integer;
            p: TProduct;
        begin
            Chart2.Title.Caption := FormatDateTime('dd.MM.YYYY HH:nn:ss',
              party.CreatedAt);
            for i := 0 to 9 do
            begin
                ListView2.Items[i].Caption := '';
            end;
            for i := 0 to current_products.Count - 1 do
            begin
                p := current_products[i];
                ListView2.Items[p.order].Caption := inttostr(p.serial);
            end;
        end);

end;

function TForm1.NodeYear(year: integer): TTreeNode;
var
    i: integer;
begin
    for i := 0 to TreeView1.Items.Count - 1 do
    begin
        if TreeView1.Items[i].Text = inttostr(year) then
        begin
            result := TreeView1.Items[i];
            exit;
        end;
    end;
end;

function TForm1.NodeYearMonth(year, month: integer): TTreeNode;
var
    i: integer;
    nore_year: TTreeNode;
begin
    nore_year := NodeYear(year);
    result := nore_year.getFirstChild;
    while result.Text <> month_number_to_str(month) do
    begin
        result := nore_year.GetNextChild(result);
    end;
end;

function TForm1.NodeYearMonthDay(year, month, day: integer): TTreeNode;
var
    i: integer;
    node_year_month: TTreeNode;
begin
    node_year_month := NodeYearMonth(year, month);
    result := node_year_month.getFirstChild;
    while result.Text <> day_to_str(day) do
    begin
        result := node_year_month.GetNextChild(result);
    end;
end;

function TForm1.NodeParty(partyID: int64): TTreeNode;
var
    i: integer;
begin
    for i := 0 to treenode_party.Keys.Count - 1 do
    begin
        result := treenode_party.Keys.ToArray[i];
        if treenode_party[result].id = partyID then
        begin
            exit;
        end;
        result := nil;
    end;
end;

function TForm1.NodeProduct(productID: int64): TTreeNode;
var
    i: integer;
begin
    for i := 0 to treenode_product.Keys.Count - 1 do
    begin
        result := treenode_product.Keys.ToArray[i];
        if treenode_product[result].id = productID then
        begin
            exit;
        end;
        result := nil;
    end;
end;

procedure TForm1.OnException(Sender: TObject; E: Exception);
begin
    terminate_error(E.Message + ' ' + E.StackTrace);
end;

// procedure TForm1.WMREGCHANGE(var Msg: TMessage);
// begin
// combobox_comports_update(ComboBox1);
// ComboBox1Change(nil);
// end;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
var
    i: integer;
    ser: TFastLineSeries;
    list_item: TListItem;
begin

    pipe_write := TPipeClient.create('UFO82_FROM_PEER_TO_MASTER');
    pipe_read := TPipeClient.create('UFO82_FROM_MASTER_TO_PEER');
    read_pipe_thread := TReadPipeThread.create;

    TreeView1Change(TreeView1, nil);
    treenode_party := TDictionary<TTreeNode, TParty>.create();
    treenode_product := TDictionary<TTreeNode, TProduct>.create();

    // Application.OnException := OnException;

    ser := TFastLineSeries.create(nil);
    ser.XValues.DateTime := true;
    ser.Title := Format('%d', [i]);
    ser.Active := true;
    Chart1.AddSeries(ser);

    ListView2.Items.Clear;
    ListView2.OnItemChecked := nil;
    for i := 1 to 10 do
    begin
        ser := TFastLineSeries.create(nil);
        ser.XValues.DateTime := true;
        ser.Title := Format('%02d', [i]);
        ser.Active := false;
        Chart2.AddSeries(ser);

        list_item := ListView2.Items.Add();

        list_item.SubItems.Add(Format('%.*d', [2, i]));
        list_item.SubItems.Add(' ');
        list_item.SubItems.Add(' ');
        list_item.SubItems.Add(' ');
        list_item.Checked := true;
        list_item.ImageIndex := -1;

    end;
    ListView2.OnItemChecked := ListView1ItemChecked;

end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
var
    wp: WINDOWPLACEMENT;
    fs: TFileStream;
begin
    fs := TFileStream.create(TPath.Combine(ExtractFilePath(paramstr(0)),
      'position.'), fmOpenWrite or fmCreate);
    if not GetWindowPlacement(Handle, wp) then
        terminate_error('GetWindowPlacement: false');
    fs.Write(wp, sizeof(wp));
    fs.Free;
end;

procedure TForm1.FormActivate(Sender: TObject);
var
    wp: WINDOWPLACEMENT;
    fs: TFileStream;
    FileName: string;
begin
    FileName := TPath.Combine(ExtractFilePath(paramstr(0)), 'position.');
    if FileExists(FileName) then
    begin
        fs := TFileStream.create(FileName, fmOpenRead);
        fs.Read(wp, sizeof(wp));
        fs.Free;
        SetWindowPlacement(Handle, wp);
    end;
    self.OnActivate := nil;
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
begin
    Button1.Enabled := (Button1.Caption <> 'Запустить опрос') or
      (ComboBox1.ItemIndex > -1);
    pipe_write.WriteInt32(MSG_PORT_NAME);
    pipe_write.WriteString(ComboBox1.Text);

end;

procedure TForm1.Button1Click(Sender: TObject);
begin
    if Button1.Caption = 'Запустить опрос' then
    begin
        pipe_write.WriteInt32(MSG_PORT_NAME);
        pipe_write.WriteString(ComboBox1.Text);
        pipe_write.WriteInt32(MSG_START_HARDWARE);
    end
    else
    begin
        pipe_write.WriteInt32(MSG_STOP_HARDWARE);
    end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
    i: integer;
begin
    for i := 0 to 9 do
    begin
        Chart2.Series[i].Clear;
    end;
    pipe_write.WriteInt32(MSG_CREATE_NEW_PARTY);
end;

procedure TForm1.TreeView1Change(Sender: TObject; Node: TTreeNode);
var
    i: integer;
begin
    if (Node = nil) or (Node.Level < 4) then
    begin
        Chart1.Visible := false;
    end
    else
    begin
        Chart1.Series[0].Clear;
        Chart1.Series[0].Active := false;
        pipe_write.WriteInt32(MSG_SENSITIVITIES_OF_PRODUCT);
        pipe_write.WriteInt64(treenode_product[Node].id);

    end;

end;

procedure TForm1.TreeView1Expanding(Sender: TObject; Node: TTreeNode;
var AllowExpansion: Boolean);
var
    n: TTreeNode;
    i: integer;
begin
    if (Node.Level = 0) then
    begin
        n := Node.getFirstChild;
        if n.Text <> '' then
            exit;
        pipe_write.WriteInt32(MSG_MONTHS_OF_YEAR);
        pipe_write.WriteInt32(StrToInt(Node.Text));
    end
    else if Node.Level = 1 then
    begin
        n := Node.getFirstChild;
        if n.Text <> '' then
            exit;
        pipe_write.WriteInt32(MSG_DAYS_OF_YEAR_MONTH);
        pipe_write.WriteInt32(StrToInt(Node.Parent.Text));
        for i := 1 to 12 do
        begin
            if month_number_to_str(i) = Node.Text then
            begin
                pipe_write.WriteInt32(i);
                exit;
            end;
        end;
        raise Exception.create('bad node: ' + Node.Text);

    end
    else if Node.Level = 2 then
    begin
        n := Node.getFirstChild;
        if n.Text <> '' then
            exit;
        pipe_write.WriteInt32(MSG_PARTIES_OF_YEAR_MONTH_DAY);

        pipe_write.WriteInt32(StrToInt(Node.Parent.Parent.Text));
        for i := 1 to 12 do
        begin
            if month_number_to_str(i) = Node.Parent.Text then
            begin
                pipe_write.WriteInt32(i);
            end;
        end;
        pipe_write.WriteInt32(StrToInt(Node.Text));
    end
    else if Node.Level = 3 then
    begin
        n := Node.getFirstChild;
        if n.Text <> '' then
            exit;
        pipe_write.WriteInt32(MSG_PRODUCTS_OF_PARTY);
        pipe_write.WriteInt64(treenode_party[Node].id);
    end

      ;

end;

procedure TForm1.ListView1CustomDrawSubItem(Sender: TCustomListView;
Item: TListItem; SubItem: integer; State: TCustomDrawState;
var DefaultDraw: Boolean);
var
    r: Trect;
    c: tcanvas;
    i, d: integer;
    ser: TChartSeries;
    AListView: TListView;
    AChart: TChart;

begin

    AListView := Sender as TListView;
    AChart := Chart1;
    if AListView = ListView2 then
        AChart := Chart2;

    DefaultDraw := false;
    c := AListView.Canvas;
    r := Item.DisplayRect(drBounds);
    for i := 0 to SubItem - 1 do
    begin
        r.left := r.left + AListView.Columns.Items[i].Width;
        r.Right := r.left + AListView.Columns.Items[i + 1].Width;
    end;
    if Item.index < 0 then
        exit;

    ser := AChart.Series[Item.index];

    if place_error[Item.index] then
    begin
        ListView2.Canvas.Font.Color := clREd;

    end
    else
    begin
        ListView2.Canvas.Font.Color := clBlack;

    end;

    if SubItem = 3 then
    begin
        d := round(r.Top + r.Height / 2);
        c.Brush.Color := ser.SeriesColor;
        c.FillRect(Rect(r.left, d - 2, r.Right, d + 2));
        SetBkMode(Sender.Canvas.Handle, TRANSPARENT);
    end
    else
    begin
        c.Font.Size := 12;
        c.Refresh;

        DefaultDraw := true;
    end;

end;

procedure TForm1.ListView1ItemChecked(Sender: TObject; Item: TListItem);
var
    AListView: TListView;
    AChart: TChart;
begin
    AListView := Sender as TListView;
    AChart := Chart1;
    if AListView = ListView2 then
        AChart := Chart2;
    AChart.Series[Item.index].Active := Item.Checked;
    pipe_write.WriteInt32(MSG_PLACE_CHECKED);
    if Item.Checked then
    begin
        pipe_write.WriteInt32(1)
    end
    else
    begin
        pipe_write.WriteInt32(0);
    end;
    pipe_write.WriteInt32(Item.index)

end;

procedure TForm1.ListView1SelectItem(Sender: TObject; Item: TListItem;
Selected: Boolean);
var
    i: integer;
    ser: TFastLineSeries;
    AListView: TListView;
    AChart: TChart;
begin
    AListView := Sender as TListView;
    AChart := Chart1;
    if AListView = ListView2 then
        AChart := Chart2;
    for i := 0 to AChart.SeriesCount - 1 do
    begin
        ser := AChart.Series[i] as TFastLineSeries;
        if i = Item.index then
        begin
            ser.LinePen.Width := 3;
        end
        else
        begin
            ser.LinePen.Width := 1;

        end;
    end;
end;

procedure TForm1.ListView2CustomDrawItem(Sender: TCustomListView;
Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
    if place_error[Item.index] then
    begin
        ListView2.Canvas.Font.Color := clREd;
    end
    else
    begin
        ListView2.Canvas.Font.Color := clBlack;
    end;

    DefaultDraw := true;
end;

procedure TForm1.ListView2Edited(Sender: TObject; Item: TListItem;
var s: string);
var
    n: integer;
begin
    n := StrToIntDef(s, 0);
    if (n > 0) or (s = '') then
    begin
        pipe_write.WriteInt32(MSG_CURRENT_PRODUCT_SERIAL);
        pipe_write.WriteInt32(Item.index);
        pipe_write.WriteInt32(n);
    end;
    s := Item.Caption;
end;

procedure TForm1.ListView2MouseDown(Sender: TObject; Button: TMouseButton;
Shift: TShiftState; X, Y: integer);
var
    Item: TListItem;
begin
    Item := ListView2.GetItemAt(X, Y);

    if (Item <> nil) and (ssCtrl in Shift) then
    begin
        Item.EditCaption;

    end;
end;

end.
