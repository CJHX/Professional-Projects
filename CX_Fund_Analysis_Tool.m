function varargout = CX_Fund_Analysis_Tool(varargin)
%CX_FUND_ANALYSIS_TOOL MATLAB code file for CX_Fund_Analysis_Tool.fig
%      CX_FUND_ANALYSIS_TOOL, by itself, creates a new CX_FUND_ANALYSIS_TOOL or raises the existing
%      singleton*.
%
%      H = CX_FUND_ANALYSIS_TOOL returns the handle to a new CX_FUND_ANALYSIS_TOOL or the handle to
%      the existing singleton*.
%
%      CX_FUND_ANALYSIS_TOOL('Property','Value',...) creates a new CX_FUND_ANALYSIS_TOOL using the
%      given property value pairs. Unrecognized properties are passed via
%      varargin to CX_Fund_Analysis_Tool_OpeningFcn.  This calling syntax produces a
%      warning when there is an existing singleton*.
%
%      CX_FUND_ANALYSIS_TOOL('CALLBACK') and CX_FUND_ANALYSIS_TOOL('CALLBACK',hObject,...) call the
%      local function named CALLBACK in CX_FUND_ANALYSIS_TOOL.M with the given input
%      arguments.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help CX_Fund_Analysis_Tool

% Last Modified by GUIDE v2.5 11-Jul-2017 22:13:06

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @CX_Fund_Analysis_Tool_OpeningFcn, ...
                   'gui_OutputFcn',  @CX_Fund_Analysis_Tool_OutputFcn, ...
                   'gui_LayoutFcn',  [], ...
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


% --- Executes just before CX_Fund_Analysis_Tool is made visible.
function CX_Fund_Analysis_Tool_OpeningFcn(hObject, eventdata, handles, varargin) %#ok<*INUSL>
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   unrecognized PropertyName/PropertyValue pairs from the
%            command line (see VARARGIN)

% Choose default command line output for CX_Fund_Analysis_Tool
handles.output = hObject;
data = struct('excelsize',[],'sheets',[],'filename',[],'pathname',[]);
set(handles.pushbutton5,'UserData',data);

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes CX_Fund_Analysis_Tool wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = CX_Fund_Analysis_Tool_OutputFcn(hObject, eventdata, handles)
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles) %#ok<*DEFNU>
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%import data from Excel in Row and column vectors (列失量） format, make sure the date are the first
%column, total asset is second and change is 3rd
%rename imported data VarName1 to 'date'
%rename imported data VarName2 to 'value'
%rename imported data VarName3 to 'r8ofreturn'

data = get(handles.pushbutton5,'UserData'); %retrieves excel data from button 5
excelsize=data.excelsize; %defines number of sheets within the excel document
sheets=data.sheets; %defines the sheet names
filename=data.filename; %defines the excel's file name
pathname=data.pathname; %defines the path to the excel file

sheetnumber=1; %starts calculations from the first sheet
presheets=sheets; %trivial variable used for the UI
postsheets=[]; %trivial variable used for the UI

for  i = 1:excelsize %runs the program for (excelsize) number of times
try
    [num,text,both]= xlsread ([pathname filename], sheetnumber); %retrieves data from excel file
    clear num; %deletes unnecessary data
    clear text; %deletes unnecessary data
    both(1,:)=[]; %retrieved data from imput excel
    inputvalue=cell2mat(both(:,2)); %converts 2nd row of both into the value data in number format
    inputvalue(isnan(inputvalue)) = []; %removes any NAN values
    inputdate = cell2mat(both(:,1)); %reads dates from input and converts them into datetime format
    inputdate(isnan(inputdate)) = []; %removes any NAN values
    clear both; %deletes retrieved data
    inputsort = 5;
catch
    inputsort = 0;
end

if inputsort > 0
    returnrate=diff(inputvalue)./inputvalue(1:end-1,:); %calculates the delta change between each day
    fullreturnrate = [zeros(1);returnrate]; %creates a return matrix used in calculating monthly profit
    d1 = datenum(inputdate(1,1)); %identifies the first day
    dlast = datenum(inputdate(end,1)); %identifies the last day
    dn=datenum(inputdate); %creates a variable with datenumbers as the dates
    [y,m,d] = datevec(dn); %identifies the Y/M/D of each date
    timetbl=[y m d]; %combines the Y/M/D above into a single table
    [ii,jj,kk]=unique(timetbl(:,1:2),'rows'); %furthermore sorts the days by months
    mthlypft=accumarray(kk,(1:numel(kk))',[],@(x) sum(fullreturnrate(x))); %calculates monthly profit of every given month
    [months,temp1] = size(mthlypft); %counts the number of months
    clear temp1; %deletes unnecessary data
    posmthrows = mthlypft(:,1) > 0; %assigns a number of 1 to positive months
    posmth=sum(posmthrows); %sum of posmthrows gives number of positive months
    mthlyposrtn=posmth/months; %calculates the monthly positive earning
    annrtn=(inputvalue(end,1)/inputvalue(1,1))^(365/(dlast-d1))-1; %calculates annual return
    annvol=std(returnrate)*sqrt(52); %calculates annual volatility
    sharp=(annrtn-0.03)/annvol; %calculates sharp index
    negrtn=returnrate; %copies return rate in preparation for calculating negative return
    rtnrows = negrtn(:,1) >= 0; %identifies positive earning days
    negrtn(rtnrows,:)=[]; %removes all positive earning days
    downrisk = std(negrtn); %calculates the downward risk
    anndownrisk=downrisk*sqrt(52);
    sortino=(annrtn-0.03)/anndownrisk;
    maxdrawdown1 = maxdrawdown(inputvalue);
    calmar = annrtn/abs(maxdrawdown1);
    basiccalc=5;
else 
    basiccalc=0;
end


try
    w=windmatlab;
    [w_wsd_data,w_wsd_codes,w_wsd_fields,w_wsd_times,w_wsd_errorid,w_wsd_reqid]=w.wsd('000016.SH,000905.SH,H11001.CSI,NH0100.NHF','close',d1-5,dlast);
    datafetch = 5;
catch
    datafetch = 0; 
end
    
    if datafetch > 0
        [refsize,wert] = size(w_wsd_times);
        clear wert;
        [datesize,zcvx]=size(inputdate);
        clear zcvx;
        tempdates = [dn; zeros((refsize-datesize),1)] ; %#ok<*AGROW>
        realdates = ismember(w_wsd_times,tempdates);
        truedata=w_wsd_data(realdates,:);
        truetimes=w_wsd_times(realdates,:);

        [givendatesize,wert] = size(inputdate);
        clear wert;
        [datadatesize,zcvx]=size(truetimes);
        clear zcvx;
        
        if givendatesize > datadatesize
            tempdates2 = truetimes;
            tempdates2 = [tempdates2; zeros((givendatesize-datadatesize),1)] ; %#ok<*AGROW>
            realdates2 = ismember(dn,tempdates2);
            inputvalue=inputvalue(realdates2,:); %#ok<*NASGU>
            inputdate=inputdate(realdates2,:);
            returnrate=fullreturnrate(realdates2,:);
            returnrate(1,:)=[];
     
            %%%%RETURN RATE IS ONE SMALLER FROM VALUE AND DATE
        end
    datasort = 5;
    else
        datasort = 0;
    end
    
    if datasort > 0
        benchrtnrate=diff(truedata)./truedata(1:end-1,:);
        sz50rtnrate=benchrtnrate(:,1);
        sz50beta= cov(returnrate,sz50rtnrate)/var(sz50rtnrate);
        sz50beta = sz50beta(1,2);
        zz500rtnrate=benchrtnrate(:,2);
        zz500beta= cov(returnrate,zz500rtnrate)/var(zz500rtnrate);
        zz500beta = zz500beta(1,2);
        zzqzrtnrate=benchrtnrate(:,3);
        zzqzbeta= cov(returnrate,zzqzrtnrate)/var(zzqzrtnrate);
        zzqzbeta = zzqzbeta(1,2);
        nhsprtnrate=benchrtnrate(:,4);
        nhspbeta= cov(returnrate,nhsprtnrate)/var(nhsprtnrate);
        nhspbeta = nhspbeta(1,2);
        dataprocess = 5;    
    else
        dataprocess = 0;
    end
    
if basiccalc > 0
    out1=[annrtn;annvol;downrisk;anndownrisk;maxdrawdown1];
    out2=[sharp;sortino;calmar];
    out3=(mthlyposrtn);
    out1 = cellfun(@(x) sprintf('%0.5f%%',x*100),num2cell(out1),'UniformOutput',false);
    out2 = cellfun(@(x) sprintf('%0.5f%',x),num2cell(out2),'UniformOutput',false);
else
    out1 = '数据错误';
    out1=cellstr(out1);
    out2 = [];
    out3 = [];
    
end

basicout= [out1;out2;out3];

if dataprocess > 0
    out4=[sz50beta;zz500beta;zzqzbeta;nhspbeta];
    out4 = cellfun(@(x) sprintf('%0.5f%%',x*100),num2cell(out4),'UniformOutput',false);
else
    out4 = out1;
end


if basiccalc > 0
    if sharp > 1
        sharpscore=5;
    else if sharp < 0
            sharpscore=0;
        else if sharp < 0.25 && sharp >= 0
                sharpscore = 1;
            else if sharp <= 0.5 && sharp > 0.25
                    sharpscore = 2;
                else if  sharp <= 0.75 && sharp > 0.5
                        sharpscore = 3;
                    else if sharp <= 1 && sharp > 0.75
                            sharpscore = 4;
                        end
                    end
                end
            end
        end
    end

    if mthlyposrtn > 0.8
        mthlyrtnscore=5;
    else if mthlyposrtn < 0.4
            mthlyrtnscore=0;
        else if mthlyposrtn < 0.5 && mthlyposrtn >= 0.4
                mthlyrtnscore = 1;
            else if mthlyposrtn < 0.6 && mthlyposrtn >= 0.5
                    mthlyrtnscore = 2;
                else if  mthlyposrtn < 0.7 && mthlyposrtn >= 0.6 %#ok<*SEPEX>
                        mthlyrtnscore = 3;
                    else if mthlyposrtn <= 0.8 && mthlyposrtn > 0.7
                            mthlyrtnscore = 4;
                        end
                    end
                end
            end
        end
    end


    if maxdrawdown1 < 0.05
        maxdrawdownscore=5;
    else if maxdrawdown1 > 0.2
            maxdrawdownscore=0;
        else if maxdrawdown1 >= 0.05 && maxdrawdown1 < 0.0875
                maxdrawdownscore = 4;
            else if maxdrawdown1 >= 0.0875 && maxdrawdown1 < 0.125
                    maxdrawdownscore = 3;
                else if  maxdrawdown1 >= 0.125 && maxdrawdown1 < 0.1625
                        maxdrawdownscore = 2;
                    else if maxdrawdown1 > 0.1625 && maxdrawdown1 <= 0.2
                            maxdrawdownscore = 1;
                        end
                    end
                end
            end
        end
    end
else
    sharpscore = out1;
    mthlyrtnscore= out1;
    maxdrawdownscore= out1;
end

    finaloutput = cell(14,3);
    
if basiccalc > 0
    finaloutput(2:10,2) = basicout;
end

if dataprocess > 0
    finaloutput(11:14,2) = out4;
end

    finaloutput{6,3} = maxdrawdownscore;
    finaloutput{7,3} = sharpscore;
    finaloutput{10,3} = mthlyrtnscore;
    finaloutput{2,1}='年化收益';
    finaloutput{3,1}='年化波动';
    finaloutput{4,1} ='下行风险';
    finaloutput{5,1} ='年化下行风险';
    finaloutput{6,1} ='最大回撤';
    finaloutput{7,1} ='夏普比率';
    finaloutput{8,1} ='所提诺比率';
    finaloutput{9,1} ='卡玛比率';
    finaloutput{10,1} ='月度正收益情况';
    finaloutput{11,1} ='β系数-上证50';
    finaloutput{12,1} ='β系数-中证500';
    finaloutput{13,1} ='β系数-中证全债';
    finaloutput{14,1} ='β系数-南华商品';
    finaloutput{1,2} ='数据';
    finaloutput{1,3} ='评分';
    
    sheetname= sheets(:,sheetnumber);
    excelend= '分析结果.xlsx';
    filenamefinal=strcat(filename,excelend);
    filenamefinal=char(filenamefinal);
    finalsheetname=char(sheetname);
    xlswrite(filenamefinal,finaloutput,finalsheetname);
    
    
    sheetnumber=sheetnumber+1;
    postsheets = [sheetname,postsheets];
    presheets(:,1)=[];
    clearvars -except sheetnumber filename pathname sheets excelsize presheets postsheets handles mthlyrtnscore;
    pause(0.05);
    set(handles.listbox4,'string',presheets);
    set(handles.listbox3,'string',postsheets);
        
end
   set(handles.listbox1,'string','计算完成');

% --- Executes on key press with focus on pushbutton1 and none of its controls.
function pushbutton1_KeyPressFcn(hObject, eventdata, handles) %#ok<*INUSD>
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  structure with the following fields (see MATLAB.UI.CONTROL.UICONTROL)
%	Key: name of the key that was pressed, in lower case
%	Character: character interpretation of the key(s) that was pressed
%	Modifier: name(s) of the modifier key(s) (i.e., control, shift) pressed
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
quit

% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

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


% --- Executes on button press in pushbutton5.
function pushbutton5_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
data = get(hObject,'UserData');

[filename, pathname] = uigetfile('*.*', 'Select the File');
[status,sheets] = xlsfinfo([pathname filename]); %#ok<*ASGLU>
clear status;
excelsize=size(sheets);
excelsize(:,1)=[];

data.filename=filename;
data.pathname=pathname;
data.excelsize=excelsize;
data.sheets=sheets;

set(handles.listbox4,'string',data.sheets);
set(hObject,'UserData',data);
set(handles.listbox1,'string','准备就绪');

% --- Executes on selection change in listbox3.
function listbox3_Callback(hObject, eventdata, handles)
% hObject    handle to listbox3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

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


% --- Executes on button press in pushbutton7.
function pushbutton7_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
msgbox('请确保日期列的格式为(月月/日日/年年年年)，从最老到最新排列。 同时，请确保文件的第一行是标题(如日期，净值等)，否则分析结果会不精确。', '重要提示','warn');


% --- If Enable == 'on', executes on mouse press in 5 pixel border.
% --- Otherwise, executes on mouse press in 5 pixel border or over listbox1.
function listbox1_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- If Enable == 'on', executes on mouse press in 5 pixel border.
% --- Otherwise, executes on mouse press in 5 pixel border or over listbox3.
function listbox3_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to listbox3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- If Enable == 'on', executes on mouse press in 5 pixel border.
% --- Otherwise, executes on mouse press in 5 pixel border or over listbox4.
function listbox4_ButtonDownFcn(hObject, eventdata, handles)
% hObject    handle to listbox4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
