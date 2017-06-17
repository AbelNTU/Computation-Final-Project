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
index = {'���d' '���~' '�콦' '��´' '�q������' '�q���q�l' '�ƾǥ���' '����' '�y��'...
         '���K' '��' '�T��' '�q�l' '���' '��B' '�[��' '����' '�T���ʳf'...
         '�ƾ�' '�ͧ�����' '�b����' '�q���P��' '���q' '�q�H����' '�q�s��' '�q�l�q��'...
         '��T�A��' '��L�q�l' '�o�q�U��' '��L' };
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
if strcmp(report_cat,'�겣�t�Ū�') == 1
    list3 = {'�y�ʸ겣','�D�y�ʸ겣','�y�ʭt��','�D�y�ʭt��'};
elseif strcmp(report_cat,'��X�l�q��') == 1
    list3 = {'��~�O��','��L��X�l�q(�b�B)'};
else
    list3 = {'��~,���,�w�ꬡ�ʤ��'};
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
%����P�Ӥ�
report_categories = cellstr(get(handles.listbox1,'String'));
report_cat = report_categories{get(handles.listbox1,'Value')};
subcatgories = cellstr(get(handles.listbox3,'String'));
subcat = subcatgories{get(handles.listbox3,'Value')};
%���~�P���q
categorys = get(handles.popupmenu1,'String'); 
target = categorys{get(handles.popupmenu1,'Value')};
companys = get(handles.popupmenu2,'String'); 
company = companys{get(handles.popupmenu2,'Value')};
num = find(strcmp(companys,company));
%�C���禬
report_e = get(handles.checkbox1,'Value');
%�C��ѻ�
report_p = get(handles.checkbox2,'Value');

switch subcat
    case '�y�ʸ겣'
        labels = {'�w�I�ڶ�','�����b�ڲb�B','��L�y�ʸ겣','��L�����ڲb�B','�s�f','�{���ά���{��'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        h = pie(handles.axes1,data(:,num));
        h1 = pie(handles.axes2,sum(data,2));
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '�D�y�ʸ겣'
        labels = {'�ƨѥX����ĸ겣�ЫD�y�ʲb�B','�ĥ��v�q�k�����b�B','���ʲ��B�t�Фγ]��',...
          '���ʤ��ʲ��b�B','�L�θ겣','��L�D�y�ʸ겣'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '�y�ʭt��'
        labels = {'�u���ɴ�','���I�u������','���I����','���I�b��','��L���I��','����ұo�|�t��','��L�y�ʭt��'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '�D�y�ʭt��'
        labels = {'�����ɴ�','�����ұo�|�t��','��L�D�y�ʭt��'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '��~�O��'
        labels = {'���P�O��','�޲z�O��','��s�o�i�O��'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    case '��L��X�l�q(�b�B)'
        labels = {'��~��B���c�]�ȳ����⤧�I���t�B','�ƨѥX����ĸ겣����{�����l�q',...
          '�{���y�q���I','�T�w�֧Q�p�e���Q�q�]�l���^','�ĥ��v�q�k�{�C�����p���~�ΦX�ꤧ��L��X�l�q�����B�X�p',...
          '��L��X�l�q'};
        data = xlsread(strcat(report_cat,'/',target,'/',subcat,'/',year,'0',season,'.xlsx'));
        pie(handles.axes1,data(:,num))
        pie(handles.axes2,sum(data,2))
        legend(handles.axes1,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
        legend(handles.axes2,labels,'Location','eastoutside','Orientation','vertical','FontSize',12);
    otherwise
        labels = categorical({'��~���ʤ��b�{���y�J�]�y�X�^','��ꬡ�ʤ��b�{���y�J�]�y�X�^','�w�ꬡ�ʤ��b�{���y�J�]�y�X�^'});
        data = xlsread(strcat(report_cat,'/',target,'/',year,'0',season,'.xlsx'));
        bar(handles.axes1,labels,data(:,num))
        bar(handles.axes2,labels,sum(data,2))
end
title(handles.axes1,company,'Color','r','FontSize',24);
title(handles.axes2,target,'Color','r','FontSize',24);
if report_e == 1
  excelname=dir(strcat('excel��/',target,'.excel/',company));
  [N,T,data]=xlsread(strcat('excel��/',target,'.excel/',company,'/',excelname(4).name));
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
  excelname=dir(strcat('excel��/',target,'.excel/',company));
  data=xlsread(strcat('excel��/',target,'.excel/',company,'/',excelname(8).name));
  
  plot(handles.axes1,data);
end
