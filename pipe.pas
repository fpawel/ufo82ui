unit pipe;

interface

type
    TPipeClient = class
    private
        { Private declarations }
        hPipe: THANDLE; // дескриптор
        buf: array [0 .. 1000] of byte;
    public
        { Public declarations }
        Constructor Create(pipe_server: string);
        procedure WriteInt32(v: LONGWORD);
        function ReadInt32: LONGWORD;

        procedure WriteString(s: string);
        function ReadString: string;

        procedure WriteInt64(value: int64);
        function ReadInt64: int64;

        procedure WriteFloat64(value: double);
        function ReadFloat64: double;

        function ReadDateTime: TDateTime;
    end;

implementation

uses sysutils, Winapi.Windows, System.WideStrUtils, system.dateutils;

type
    _LONGWORD_BYTES = record

        case Integer of
            0:
                (bytes: array [0 .. 3] of byte);
            1:
                (value: LONGWORD);
    end;

type
    _INT64_BYTES = record

        case Integer of
            0:
                (bytes: array [0 .. 7] of byte);
            1:
                (value_int64: int64);

            2:
                (value_double: Double);
    end;

procedure terminate_error(error_text: string);
var
    f: TextFile;
begin
    AssignFile(f, ExtractFileDir(paramstr(0)) + '\pipe_fail.txt');
    ReWrite(f);
    WriteLn(f, error_text);
    CloseFile(f);
    ExitProcess(1);
end;

function millis_to_datetime(x_millis: int64): TDateTime;
begin
    result := IncHour(IncMilliSecond(EncodeDateTime(1970, 1, 1, 0, 0, 0, 0),
      x_millis), 3);
end;

Constructor TPipeClient.Create(pipe_server: string);
begin
    hPipe := CreateFileW(PWideChar('\\.\pipe\$' + pipe_server + '$'),
      GENERIC_READ or GENERIC_WRITE, FILE_SHARE_READ or FILE_SHARE_WRITE, nil,
      OPEN_EXISTING, 0, 0);
    if hPipe = INVALID_HANDLE_VALUE then
        terminate_error('hPipe = INVALID_HANDLE_VALUE');

end;


function TPipeClient.ReadFloat64: double;
var
    x: _INT64_BYTES;
    readed_count: DWORD;
begin
    if not ReadFile(hPipe, x.bytes, 8, readed_count, nil) then
    begin
        terminate_error('ReadFloat64: ReadFile error');
    end;
    if readed_count <> 8 then
    begin
        terminate_error('ReadFloat64: readed_count <> 8');
    end;
    result := x.value_double;

end;

procedure TPipeClient.WriteFloat64(value: double);
var
    writen_count: DWORD;
    x: _INT64_BYTES;

begin
    x.value_double := value;
    if not(WriteFile(hPipe, x.bytes, 8, writen_count, nil)) then
    begin
        terminate_error('WriteFloat64: WriteFile error');
    end;
    if writen_count <> 8 then
    begin
        terminate_error('WriteFloat64: writen_count <>8');
    end;
end;

function TPipeClient.ReadInt64: int64;
var
    x: _INT64_BYTES;
    readed_count: DWORD;
begin
    if not ReadFile(hPipe, x.bytes, 8, readed_count, nil) then
    begin
        terminate_error('ReadInt64: ReadFile error');
    end;
    if readed_count <> 8 then
    begin
        terminate_error('ReadInt64: readed_count <> 8: ' + inttostr(readed_count));
    end;
    result := x.value_int64;

end;

procedure TPipeClient.WriteInt64(value: int64);
var
    writen_count: DWORD;
    x: _INT64_BYTES;

begin
    x.value_int64 := value;
    if not(WriteFile(hPipe, x.bytes, 8, writen_count, nil)) then
    begin
        terminate_error('WriteInt64: WriteFile error');
    end;
    if writen_count <> 8 then
    begin
        terminate_error('WriteInt64: writen_count <>8');
    end;
end;

function TPipeClient.ReadInt32: LONGWORD;
var
    x: _LONGWORD_BYTES;
    readed_count: DWORD;
begin
    if not ReadFile(hPipe, x.bytes, 4, readed_count, nil) then
    begin
        terminate_error('ReadInt32: ReadFile error');
    end;
    if readed_count <> 4 then
    begin
        terminate_error('ReadInt32: readed_count <> 4:' + inttostr(readed_count));
    end;
    result := x.value;
end;

procedure TPipeClient.WriteInt32(v: LONGWORD);
var
    writen_count: DWORD;
    x: _LONGWORD_BYTES;

begin
    x.value := v;
    if not(WriteFile(hPipe, x.bytes, 4, writen_count, nil)) then
    begin
        terminate_error('WriteInt: WriteFile error');
    end;
    if writen_count <> 4 then
    begin
        terminate_error('WriteInt: writen_count <> 4');
    end;
end;

procedure TPipeClient.WriteString(s: string);
var
    writen_count: DWORD;
    i, len: Cardinal;

begin
    len := Length(s) * sizeof(char);
    if len >= Length(self.buf) + 1 then
    begin
        raise Exception.Create('Превышен размер буфера');
    end;
    self.WriteInt32(len);
    if len = 0 then
        exit;

    Move(s[1], self.buf, len);

    if not(WriteFile(hPipe, self.buf, len, writen_count, nil)) then
    begin
        terminate_error('WriteString: WriteFile error');
    end;
    if writen_count <> len then
    begin
        terminate_error('WriteString: writen_count <> 1');
    end;
end;

function TPipeClient.ReadString: string;
var
    readed_count: DWORD;
    len: LONGWORD;
    b: array [0 .. 1000] of byte;
begin
    len := ReadInt32;
    if len = 0 then
    begin
        result := '';
        exit;
    end;
    if not ReadFile(hPipe, b, len, readed_count, nil) then
    begin
        terminate_error('ReadString: ReadFile error');
    end;
    if readed_count <> len then
    begin
        terminate_error(Format('ReadString: readed_count = %d <> str len = %d',
          [readed_count, len]));
    end;
    SetString(result, PAnsiChar(@b[0]), len);
    Finalize(b);

end;

function TPipeClient.ReadDateTime: TDateTime;
var s: string;
begin
    result := millis_to_datetime(ReadInt64);
    s := TimeToStr(result);
    s := '';

end;


end.
