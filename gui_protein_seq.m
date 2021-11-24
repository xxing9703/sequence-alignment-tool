function varargout = gui_protein_seq(varargin)
% GUI_PROTEIN_SEQ MATLAB code for gui_protein_seq.fig
%      GUI_PROTEIN_SEQ, by itself, creates a new GUI_PROTEIN_SEQ or raises the existing
%      singleton*.
%
%      H = GUI_PROTEIN_SEQ returns the handle to a new GUI_PROTEIN_SEQ or the handle to
%      the existing singleton*.
%
%      GUI_PROTEIN_SEQ('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in GUI_PROTEIN_SEQ.M with the given input arguments.
%
%      GUI_PROTEIN_SEQ('Property','Value',...) creates a new GUI_PROTEIN_SEQ or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before gui_protein_seq_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to gui_protein_seq_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help gui_protein_seq

% Last Modified by GUIDE v2.5 20-Jan-2018 00:06:49

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @gui_protein_seq_OpeningFcn, ...
                   'gui_OutputFcn',  @gui_protein_seq_OutputFcn, ...
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


% --- Executes just before gui_protein_seq is made visible.
function gui_protein_seq_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to gui_protein_seq (see VARARGIN)

% Choose default command line output for gui_protein_seq
handles.output = hObject;
set(handles.status,'string','Ready');
 set(handles.status,'BackgroundColor','c');
% Update handles structure
guidata(hObject, handles);

% UIWAIT makes gui_protein_seq wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = gui_protein_seq_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in bt_load.
function bt_load_Callback(hObject, eventdata, handles)
% hObject    handle to bt_load (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
[filename, pathname] = ...
     uigetfile('*.xlsx','File Selector','MultiSelect','off');
 set(handles.status,'string','Loading......');
 set(handles.status,'BackgroundColor','r');
 drawnow();
 
 if filename==0
 else
 fname=fullfile(pathname,filename); 
 %dt=importdata(fname);
 [~,~,dt]=xlsread(fname);
 dt = cellfun(@num2str,dt, 'UniformOutput',false);
 sz=size(dt,1);
 start=0; % no header
 set(handles.uitable1,'data',dt(1+start:end,:));
 end
 
 set(handles.status,'string','Ready');
 set(handles.status,'BackgroundColor','c');
 

% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)
% hObject    handle to checkbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of checkbox1


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

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


% --- Executes on button press in bt_run.
function bt_run_Callback(hObject, eventdata, handles)
% hObject    handle to bt_run (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
set(handles.status,'string','Converting......'); 
set(handles.status,'BackgroundColor','r');
drawnow();
dt=get(handles.uitable1,'data');
 base=dt{1,3};  %base DNA sequence
 cd=get(handles.popupmenu1,'value');
 
 codon=importdata('codon.xlsx');
set(handles.status,'string','Codon');  
 switch cd
     case 1
 codon_item = codon(:,[1,2]);
     case 2
 codon_item = codon(:,[1,3]);
     case 3
 codon_item = codon(:,[1,4]);
 end 
 
  sz=size(dt,1);
 for i=1:sz
    P(i).Header=dt{i,1};
    P(i).Sequence=dt{i,2};
   
 end
 dist = seqpdist(P,'ScoringMatrix','GONNET');
 tree = seqlinkage(dist,'average',P);
 ma = multialign(P,tree,'ScoringMatrix',...
                {'pam150','pam200','pam250'});
 %showalignment(ma,'SimilarColor','b')
for i=1:sz
    out{i,1}=ma(i).Header;
    out{i,2}=ma(i).Sequence;
    [~,~,s0]=p2d(ma(i).Sequence,codon_item);
    out{i,3}=s0;
end
out{1,3}=stretch(ma(1).Sequence,base);

set(handles.status,'string','Out');  
len=length(ma(1).Sequence);

for i=2:sz
    for j=1:len
      if strcmp(ma(i).Sequence(j),ma(1).Sequence(j))
          out{i,3}(j*3-2:j*3)=out{1,3}(j*3-2:j*3);
      end      
    end
end


set(handles.uitable2,'data',out);
set(handles.status,'string','Ready');
set(handles.status,'BackgroundColor','c');


% --- Executes on button press in bt_save.
function bt_save_Callback(hObject, eventdata, handles)
% hObject    handle to bt_save (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

dt=get(handles.uitable2,'data');  %all protein sequences
sz=size(dt,1);
len=size(dt{1,2},2);
skip=2;
for i=1:len
 W{1,i+1}=num2str(i);  %W row 1
end

for k=1:sz
  W{k*skip+1,1}=dt{k,1}; %header
    
  for i=1:len
      char=dt{k,2}; %Protein sequence
      W{k*skip+1,i+1}=char(i);  %W row (2:end) Protein seq.
      if skip==2
       char2=dt{k,3};
       W{k*skip+2,i+1}=char2(i*3-2:i*3);  %W row (2:end) DNA seq  
      end
  end
      
end
[file,path] = uiputfile('*.xlsx','Save to Excel');
set(handles.status,'string','Exporting......');
set(handles.status,'BackgroundColor','r');
drawnow();

savefile=fullfile(path,file);
xlswrite(savefile,dt,'sheet2');
xlswrite(savefile,W,'sheet1');


green='A1';red='A1';yellow='A1';

%find the cells, mark the color
Excel = actxserver('excel.application');
%WB = Excel.Workbooks.Open(fullfile(pwd, savefile),0,false);
WB = Excel.Workbooks.Open(savefile,0,false);

%WB.Worksheets.Item(1).Range(rg).Interior.ColorIndex = 4;
mx=245;
for i=1:len 
   for j=1:sz
    rg=strcat(ExcelCol(i+1),int2str(j*skip+1));  
    rg=rg{1};
       if strcmp(dt{j,2}(i),dt{1,2}(i)) 
           green=[green,',',rg];
           if length(green)>mx
            %WB.Worksheets.Item(1).Range(rg).Interior.ColorIndex = 4;
            WB.Worksheets.Item(1).Range(green).Interior.ColorIndex = 4;
            green='A1';
           end
       else 
           red=[red,',',rg];
           if length(red)>mx
           %WB.Worksheets.Item(1).Range(rg).Interior.ColorIndex = 3;
           WB.Worksheets.Item(1).Range(red).Interior.ColorIndex = 3;
           red='A1';
           end
       end
       if strcmp(dt{j,2}(i),'-') 
           yellow=[yellow,',',rg];
           if length(yellow)>mx
           %WB.Worksheets.Item(1).Range(rg).Interior.ColorIndex = 6; 
           WB.Worksheets.Item(1).Range(yellow).Interior.ColorIndex = 6;
           yellow='A1';
           end
       end
      %WB.Save();
    end
end
WB.Worksheets.Item(1).Range(green).Interior.ColorIndex = 4;
WB.Worksheets.Item(1).Range(red).Interior.ColorIndex = 3;
WB.Worksheets.Item(1).Range(yellow).Interior.ColorIndex = 6;
WB.Save();

WB.Close();
Excel.Quit();
set(handles.status,'string','Ready');
set(handles.status,'BackgroundColor','c');
fprintf('done\n');
