unit Unit1;

interface

uses
  ActiveX, //For List of Camera devices

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    ComboBox1: TComboBox;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }

  public
    { Public declarations }
    procedure listwindowscam(xcombobox:TComboBox);

  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}
//////////////////////////BEGIN GET A LIST OF VIDEO DEVICES
type
ICreateDevEnum = interface(IUnknown)
['{29840822-5B84-11D0-BD3B-00A0C911CE86}']
function CreateClassEnumerator(const clsidDeviceClass: TGUID;
out ppEnumMoniker: IEnumMoniker; dwFlags: DWORD): HRESULT; stdcall;
end;
procedure TForm1.listwindowscam(xcombobox:TComboBox);
const
CLSID_SystemDeviceEnum:         TGUID = (D1:$62BE5D10;D2:$60EB;D3:$11D0;D4:($BD,$3B,$00,$A0,$C9,$11,$CE,$86));
IID_ICreateDevEnum:             TGUID = '{29840822-5B84-11D0-BD3B-00A0C911CE86}';
CLSID_VideoInputDeviceCategory: TGUID = (D1:$860BB310;D2:$5D01;D3:$11D0;D4:($BD,$3B,$00,$A0,$C9,$11,$CE,$86));
var
pDevEnum     : ICreateDevEnum;
pClassEnum   : IEnumMoniker;
pMoniker     : IMoniker;
pPropertyBag : IPropertyBag;
v            : OleVariant;
cFetched     : ulong;
xname        : string;
begin
try
CoCreateInstance (CLSID_SystemDeviceEnum,nil,CLSCtx_INPROC,IID_ICreateDevEnum,pDevEnum);
pClassEnum := nil;
pDevEnum.CreateClassEnumerator (CLSID_VideoInputDeviceCategory, pClassEnum, 0);
VariantInit(v);
pMoniker := nil;
while (pClassEnum.Next(1, pMoniker, @cFetched) = S_OK) do
begin
pPropertyBag := nil;
if S_OK = pMoniker.BindToStorage(nil, nil, IPropertyBag, pPropertyBag) then
begin
if S_OK = pPropertyBag.Read('FriendlyName', v, nil) then
begin
xname := v;
xcombobox.Style := csDropDownList;
xcombobox.Items.Add(xname);
xcombobox.ItemIndex := 0;
end;
end;
VariantClear(v);
end;
pClassEnum := nil;
pMoniker := nil;
pDevEnum := nil;
pPropertyBag := nil;
except
ShowMessage('Warning!...No Camera Available!');
end;
end;
//////////////////////////END GET A LIST OF VIDEO DEVICES

procedure TForm1.FormCreate(Sender: TObject);
begin
//
listwindowscam(combobox1);
end;

end.
