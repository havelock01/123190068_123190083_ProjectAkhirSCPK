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

% Last Modified by GUIDE v2.5 03-Jul-2021 21:05:26

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




% --- Executes on button press in pb_skor.
function pb_skor_Callback(hObject, eventdata, handles)
optsData = detectImportOptions('Data_kasar.xlsx','DataRange', 'B2:I43');
readData = readmatrix('Data_kasar.xlsx', optsData);
optsNama = detectImportOptions('Data_kasar.xlsx','DataRange', 'A2:A43');
readNama = readmatrix('Data_kasar.xlsx', optsNama);

tableNama = array2table(readNama);
tableData = array2table(readData);

dataAll = table2cell([tableNama tableData]);
set(handles.uitable1,'data',dataAll);
% hObject    handle to pb_skor (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)



function tf_jarak_Callback(hObject, eventdata, handles)
% hObject    handle to tf_jarak (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_jarak as text
%        str2double(get(hObject,'String')) returns contents of tf_jarak as a double


% --- Executes during object creation, after setting all properties.
function tf_jarak_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_jarak (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function tf_harga_Callback(hObject, eventdata, handles)
% hObject    handle to tf_harga (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_harga as text
%        str2double(get(hObject,'String')) returns contents of tf_harga as a double


% --- Executes during object creation, after setting all properties.
function tf_harga_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_harga (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function tf_pelayanan_Callback(hObject, eventdata, handles)
% hObject    handle to tf_pelayanan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_pelayanan as text
%        str2double(get(hObject,'String')) returns contents of tf_pelayanan as a double


% --- Executes during object creation, after setting all properties.
function tf_pelayanan_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_pelayanan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function tf_rasa_Callback(hObject, eventdata, handles)
% hObject    handle to tf_rasa (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_rasa as text
%        str2double(get(hObject,'String')) returns contents of tf_rasa as a double


% --- Executes during object creation, after setting all properties.
function tf_rasa_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_rasa (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)

kjarak=str2double(get(handles.tf_jarak,'String'));
kpelayanan=str2double(get(handles.tf_pelayanan,'String'));
kfasilitas=str2double(get(handles.tf_fasilitas,'String'));
ksuasana=str2double(get(handles.tf_suasana,'String'));
kharga=str2double(get(handles.tf_harga,'String'));
krasa=str2double(get(handles.tf_rasa,'String'));
karoma=str2double(get(handles.tf_aroma,'String'));
ktampilan=str2double(get(handles.tf_tampilan,'String'));

w=[kjarak kpelayanan kfasilitas ksuasana kharga krasa karoma ktampilan];
k=[0 1 1 1 0 1 1 1];

opts = detectImportOptions('WP.xlsx');
opts.SelectedVariableNames = [2:9];
data = readmatrix('WP.xlsx',opts);

[m n]=size (data);
w=w./sum(w);

for j=1:n,
    if k(j)==0, w(j)=-1*w(j);
    end;
end;
for i=1:m,
    S(i)=prod(data(i,:).^w);
end;

V= S/sum(S);
rank = sort(V,'descend');

for i=1:m
    hasil(i) = rank(i);
end

opts2 = detectImportOptions('WP.xlsx');
opts2.SelectedVariableNames = [1]; 

namaBrand = readmatrix('WP.xlsx',opts2);

for i=1:m
 for j=1:m
   if(hasil(i) == V(j))
    sorting(i) = namaBrand(j);
    break
   end
 end
end

sorting = sorting';

set(handles.uitable2,'data',sorting);
% hObject    handle to pushbutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)



function tf_fasilitas_Callback(hObject, eventdata, handles)
% hObject    handle to tf_fasilitas (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_fasilitas as text
%        str2double(get(hObject,'String')) returns contents of tf_fasilitas as a double


% --- Executes during object creation, after setting all properties.
function tf_fasilitas_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_fasilitas (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

function tf_suasana_Callback(hObject, eventdata, handles)
% hObject    handle to tf_suasana (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_suasana as text
%        str2double(get(hObject,'String')) returns contents of tf_suasana as a double


% --- Executes during object creation, after setting all properties.
function tf_suasana_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_suasana (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function tf_aroma_Callback(hObject, eventdata, handles)
% hObject    handle to tf_aroma (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_aroma as text
%        str2double(get(hObject,'String')) returns contents of tf_aroma as a double


% --- Executes during object creation, after setting all properties.
function tf_aroma_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_aroma (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function tf_tampilan_Callback(hObject, eventdata, handles)
% hObject    handle to tf_tampilan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of tf_tampilan as text
%        str2double(get(hObject,'String')) returns contents of tf_tampilan as a double


% --- Executes during object creation, after setting all properties.
function tf_tampilan_CreateFcn(hObject, eventdata, handles)
% hObject    handle to tf_tampilan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
