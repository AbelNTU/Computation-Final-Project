function varargout = GUI(varargin)
% GUI MATLAB code for GUI.fig
%      GUI, by itself, creates a new GUI or raises the existing
%      singleton*.
%
%      H = GUI returns the handle to a new GUI or the handle to
%      the existing singleton*.
%
%      GUI('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in GUI.M with the given input arguments.
%
%      GUI('Property','Value',...) creates a new GUI or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before GUI_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to GUI_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help GUI

% Last Modified by GUIDE v2.5 17-Jun-2017 22:05:10

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @GUI_OpeningFcn, ...
                   'gui_OutputFcn',  @GUI_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before GUI is made visible.
function GUI_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to GUI (see VARARGIN)

% Choose default command line output for GUI
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes GUI wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = GUI_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
categorys = get(handles.popupmenu1,'String'); 
target = categorys{get(handles.popupmenu1,'Value')};
index = {'水泥' '食品' '塑膠' '紡織' '電機機械' '電器電纜' '化學生醫' '玻璃' '造紙'...
         '鋼鐵' '橡膠' '汽車' '電子' '營建' '航運' '觀光' '金融' '貿易百貨'...
         '化學' '生技醫療' '半導體' '電腦周邊' '光電' '通信網路' '電零組' '電子通路'...
         '資訊服務' '其他電子' '油電燃氣' '其他' };
category_num = find(strcmp(target,index));
data = xlsread('Desktop/index.xlsx');
dic = data(category_num+1,:);
dic = num2cell(dic(~isnan(dic)));
set(handles.popupmenu2,'String',dic);

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu1


% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu2

% --- Executes during object creation, after setting all properties.
function popupmenu2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hint: get(hObject,'Value') returns toggle state of checkbox1


% --- Executes on button press in checkbox2.
function checkbox2_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hint: get(hObject,'Value') returns toggle state of checkbox2


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
report_categories = cellstr(get(hObject,'String'));
report_cat = report_categories{get(hObject,'Value')};
if strcmp(report_cat,'資產負債表') == 1
    list3 = {'流動資產','非流動資產','流動負債','非流動負債'};
elseif strcmp(report_cat,'綜合損益表') == 1
    list3 = {'營業費用','其他綜合損益(淨額)'};
else
    list3 = {'營業,投資,籌資活動比較'};
end
set(handles.listbox3,'String',list3);
% Hints: contents = cellstr(get(hObject,'String')) returns listbox1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox1


% --- Executes during object creation, after setting all properties.
function listbox1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in listbox2.
function listbox2_Callback(hObject, eventdata, handles)
% hObject    handle to listbox2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
years = cellstr(get(hObject,'String'));
year = years{get(hObject,'Value')};
% Hints: contents = cellstr(get(hObject,'String')) returns listbox2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox2


% --- Executes during object creation, after setting all properties.
function listbox2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on selection change in listbox4.
function listbox4_Callback(hObject, eventdata, handles)
% hObject    handle to listbox4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns listbox4 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox4


% --- Executes during object creation, after setting all properties.
function listbox4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on selection change in listbox3.
function listbox3_Callback(hObject, eventdata, handles)
% hObject    handle to listbox3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
subcatgories = cellstr(get(hObject,'String'));
subcat = subcatgories{get(hObject,'Value')};

% Hints: contents = cellstr(get(hObject,'String')) returns listbox3 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox3


% --- Executes during object creation, after setting all properties.
function listbox3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
years = get(handles.listbox2,'String'); 
year = years{get(handles.listbox2,'Value')};
seasons = get(handles.listbox4,'String'); 
season = seasons{get(handles.listbox4,'Value')};
%報表與細分
report_categories = cellstr(get(handles.listbox1,'String'));
report_cat = report_categories{get(handles.listbox1,'Value')};
subcatgories = cellstr(get(handles.listbox3,'String'));
subcat = subcatgories{get(handles.listbox3,'Value')};
%產業與公司
categorys = get(handles.popupmenu1,'String'); 
target = categorys{get(handles.popupmenu1,'Value')};
companys = get(handles.popupmenu2,'String'); 
company = companys{get(handles.popupmenu2,'Value')};
num = find(strcmp(companys,company));
%每月營收
report_e = get(handles.checkbox1,'Value');
%每月股價
report_p = get(handles.checkbox2,'Value');

switch subcat
    case '流動資產'
        labels = {'預付款項','應收帳款淨額','其他流動資產','其他應收款淨額','存貨','現金及約當現金'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        h = pie(handles.axes1,data(:,num));
        h1 = pie(handles.axes2,sum(data,2));
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '非流動資產'
        labels = {'備供出售金融資產－非流動淨額','採用權益法之投資淨額','不動產、廠房及設備',...
          '投資性不動產淨額','無形資產','其他非流動資產'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '流動負債'
        labels = {'短期借款','應付短期票券','應付票據','應付帳款','其他應付款','當期所得稅負債','其他流動負債'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '非流動負債'
        labels = {'長期借款','遞延所得稅負債','其他非流動負債'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '營業費用'
        labels = {'推銷費用','管理費用','研究發展費用'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '其他綜合損益(淨額)'
        labels = {'國外營運機構財務報表換算之兌換差額','備供出售金融資產未實現評價損益',...
          '現金流量避險','確定福利計畫精算利益（損失）','採用權益法認列之關聯企業及合資之其他綜合損益之份額合計',...
          '其他綜合損益'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    otherwise
        labels = categorical({'營業活動之淨現金流入（流出）','投資活動之淨現金流入（流出）','籌資活動之淨現金流入（流出）'});
        data = xlsread(strcat(report_cat,'/',target,'/',year,'0',season,'.xlsx'));
        bar(handles.axes1,labels,data(:,num))
        bar(handles.axes2,labels,sum(data,2))
end
title(handles.axes1,company,'Color','r','FontSize',24);
title(handles.axes2,target,'Color','r','FontSize',24);
if report_e == 1
  excelname=dir(strcat('excel檔/',target,'.excel/',company));
  [N,T,data]=xlsread(strcat('excel檔/',target,'.excel/',company,'/',excelname(4).name));
  [e,f]=size(T);
  f=f-1;
  x = zeros(f);
      for i=1:f
          x(i)=i-1;
          y(i)=N(2,i);
      end
   plot(handles.axes2,x,y);
end    
if report_p == 1    
  excelname=dir(strcat('excel檔/',target,'.excel/',company));
  data=xlsread(strcat('excel檔/',target,'.excel/',company,'/',excelname(8).name));
  
  plot(handles.axes1,data);
end
