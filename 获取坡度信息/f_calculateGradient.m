function  f_calculateGradient( r_fileName )
%UNTITLED Summary of this function goes here
%   Detailed explanation goes here
samplingTime = 5; %%����ʱ��
path = pwd;
%% �����ʾ��Ϣ
disp(strcat('���ڴ���',r_fileName));
disp('���Ժ�......');
%% ����Excel���
[excelData,str] = xlsread(r_fileName,1);               %��ȡԭʼ���ݱ��е����ݣ�strΪ���ݱ��е��ַ���dataΪ���ݱ��е�����
[excelRow,excelColumn] = size(excelData);        %%��ȡ���ݱ��е����и���
value =  zeros(excelRow,4);                      %����һ����Ӧ������1�еľ������ڴ洢����������
invalidDataNum = zeros(1,4);                     %��¼���ݱ�ǰ����Ч���ݵĸ�����
[m,n] = size(str);                              %% ���ݱ����ַ��ĸ���
needStr = {'����','�ۼ����','GPS����','GPS���','GPS����'}; %% �����¶���Ҫ��������
needStrStationIn_value = zeros(1,5);                        %% ����������ԭʼ���ݱ��е�λ��

%% �ҳ���Ҫ����������ԭʼ���ݱ��е�λ��
for i = 1 :n                        
    for j = 1: 5
        if strcmp(str(1,i),needStr(1,j))>0
            needStrStationIn_value(1,j) = i-1;      %% -1����Ϊ��ԭ���ݱ��е�һ��Ϊʱ�䣬MATLAB��ȡ�����ݾ�����û����һ�С�excelData��������������ԭʼexcel��һ�У���һ�С�
        end
    end
end
format short g                                      %%������ʾ��ʽ
%% �Ǳ��ټ����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    speedSum =  excelData(row_x+1,needStrStationIn_value(1,1)) + excelData(row_x,needStrStationIn_value(1,1));
    if speedSum == 0                                                        % speedSum=0ʱ����ĸΪ0����Ч����
        if invalidDataNum(1,1) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,1) = 0;
        else
            value(row_x,1) = value(row_x-1,1);                          %��¼��Ч���ݸ���������һ�����ݽ������
        end
    else                                                                   %��Ч����
        if invalidDataNum(1,1) == 0                                        
            invalidDataNum(1,1) = row_x;                                  %��û�м�¼��Ч����ʱ��¼����Ч���ݵĸ���
        end     
        mid_value_2 = asind(gpsElevationDiffe/(speedSum/2*samplingTime/3600*1000)); 
        if isreal(mid_value_2)                                              
            podu = tand(mid_value_2);
            value(row_x,1) = podu *100;                                       %д�������
        else
            value(row_x,1) = value(row_x-1,1);                         %%���ָ���ʱ��Ϊ���ݳ���ʹ����һ���������
        end
    end
end
%% �ۼ���̼����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    accumulativeMileageDiffe =  excelData(row_x+1,needStrStationIn_value(1,2)) - excelData(row_x,needStrStationIn_value(1,2));
    if accumulativeMileageDiffe == 0                                        %%�ۼ���̲����0���������ʱʹ����һ�е����ݽ������
        if invalidDataNum(1,2) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,2) = 0;
        else
            value(row_x,2) = value(row_x-1,2);
        end
    else   
        if invalidDataNum(1,2) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            invalidDataNum(1,2) = row_x;
        end
        mid_value_2 = asind(gpsElevationDiffe/accumulativeMileageDiffe/1000); 
            podu =  tand(mid_value_2);
            value(row_x,2) = podu*100;                                         %д�������
    end
end
%% GPS���ټ����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    gpsSpeedSum =  excelData(row_x+1,needStrStationIn_value(1,3)) + excelData(row_x,needStrStationIn_value(1,3));
    if gpsSpeedSum == 0                                                    %%�ٶȻ����0���������ʱʹ����һ�е����ݽ������
        if invalidDataNum(1,3) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,3) = 0;
        else
            value(row_x,3) = value(row_x-1,3);
        end
    else   
        if invalidDataNum(1,3) == 0                                         % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            invalidDataNum(1,3) = row_x;
        end
        mid_value_2 = asind(gpsElevationDiffe/(gpsSpeedSum/2*samplingTime/3600*1000)); %% ע���޸Ĳ���ʱ��
        if isreal(mid_value_2)                                              %%���ָ���ʱ��Ϊ���ݳ���ʹ����һ���������
            podu =  tand(mid_value_2);
            value(row_x,3) = podu*100;                                        %д�������
        else
            value(row_x,3) = value(row_x-1,3);
        end
    end
end
%% GPS��̼����¶�
for row_x = 1: excelRow - 1
    gpsElevationDiffe = excelData(row_x+1,needStrStationIn_value(1,5)) - excelData(row_x,needStrStationIn_value(1,5));
    gpsMileageDiffe =  excelData(row_x+1,needStrStationIn_value(1,4)) - excelData(row_x,needStrStationIn_value(1,4));
    if gpsMileageDiffe == 0                                                %%�ٶȻ����0���������ʱʹ����һ�е����ݽ������
        if invalidDataNum(1,4) == 0                                        % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            value(row_x,4) = 0;
        else
            value(row_x,4) = value(row_x-1,4);
        end
    else   
        if invalidDataNum(1,4) == 0                                        % ��û�м�¼��Ч���ݸ���ʱ��Ч���ݵ�λ�����0
            invalidDataNum(1,4) = row_x;
        end
        mid_value_2 = asind(gpsElevationDiffe/gpsMileageDiffe/1000); 
            podu = tand(mid_value_2);
            value(row_x,4) = podu*100;                                        %д�������
    end
end
%% �˲��㷨--ȥ��>0.2��<-0.2�����ݣ�����һ���������
for i = 1:4
    for j = 2:excelRow
        if value(j,i)>20||value(j,i)<-20
            value(j,i) = value(j-1,i);
        end
    end
end
i = find('.'==r_fileName);
imname = r_fileName(1:i-1); %% imnameΪ������׺�ļ����� 
outFile = strcat(imname,'_output');
if exist(outFile)   %% �������output�ļ��У���ɾ��
     rmdir (outFile,'s');
end
mkdir(outFile);%% ����һ��Output�ļ���
cd(fullfile(path,outFile));       %%����outputĿ¼
poduFile = strcat(imname,'_podu.xls'); %%��ɴ�excle�ļ�����podu�ļ���
value_2 = value(max(invalidDataNum(:)):excelRow,1:4);                               %%ȡ����������Ч���ݣ�������Ч����
colname={'���','ʱ��','�Ǳ��ټ����¶�','�ۼ���̼����¶�','GPS���ټ����¶�','GPS��̼����¶�'};    %%����ÿһ�е���������
warning off MATLAB:xlswrite:AddSheet;   %%��ֹ����warning���� 
xlswrite(poduFile, colname, 'sheet1','A1');
xuhao = linspace(1,m-max(invalidDataNum(:)),m-max(invalidDataNum(:)));
xlswrite(poduFile, xuhao', 'sheet1','A2');                %%���
xlswrite(poduFile,str(max(invalidDataNum(:))+1:m,1), 'sheet1','B2');              %%ʱ��
xlswrite(poduFile,value_2, 'sheet1','C2');                    %%����������
%% ��������Ҫ�����ݿ���һ�ݷ���Sheet2�У��Ա��ֶ�����ʱʹ�á�
xlswrite(poduFile,str(:,1), 'sheet2','A1');
xlswrite(poduFile,needStr, 'sheet2','B1');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,1)), 'sheet2','B2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,2)), 'sheet2','C2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,3)), 'sheet2','D2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,4)), 'sheet2','E2');
xlswrite(poduFile,excelData(:,needStrStationIn_value(1,5)), 'sheet2','F2');
colname2={'=TAN(ASIN((F3-F2)/((B2+B3)/2*10/3600*1000)))','=TAN(ASIN((F3-F2)/(C3-C2)/1000))','=TAN(ASIN((F3-F2)/((D3+D2)/2*10/3600*1000)))','=TAN(ASIN((F3-F2)/(E3-E2)/1000))'};
xlswrite(poduFile,colname2, 'sheet2','G2');

%% Sheet������
path = pwd;
filePath = fullfile(path,poduFile);
e = actxserver('Excel.Application');
ewb = e.workbooks.Open(filePath);
ewb.Worksheets.Item(1).Name = '������¶�';
ewb.Worksheets.Item(2).Name = '�����¶�ʹ�õ�����';
ewb.Save 
ewb.Close(false)
e.Quit

%% ��������ͼ
fh=figure;
set(fh,'visible','off');    %% ��������ʱ����ʾfigure����
plot(xuhao',value_2(:,1),'y-');
hold on;
plot(xuhao',value_2(:,2),'b-');
hold on;
plot(xuhao',value_2(:,3),'g-');
hold on;
plot(xuhao',value_2(:,4),'r-');
hold on;
grid on;%%��ʾ������
legend('�Ǳ��ټ����¶�','�ۼ���̼����¶�','GPS���ټ����¶�','GPS��̼����¶�');
titleFile = strcat(imname,'�����¶�'); %%��ɴ�excle�ļ�����podu�ļ���
title(titleFile);
xlabel('ʱ��˳��');
ylabel('�¶�');

%% ������ͼ����EXCE��
print -dbitmap
spwd = [pwd '\'];
filespec_user=[spwd poduFile];
try
Excel=actxGetRunningServer('Excel.Application');
catch
Excel = actxserver('Excel.Application');
end;
%����Excel����Ϊ�ɼ�
set(Excel, 'Visible', 0);
Workbooks = Excel.Workbooks;
%�������ļ����ڣ��򿪸ò����ļ��������½�һ���������������棬�ļ���Ϊ����.Excel
if exist(filespec_user,'file');
Workbook = invoke(Workbooks,'Open',filespec_user);
else
Workbook = invoke(Workbooks, 'Add');
Workbook.SaveAs(filespec_user);
end
%���ع�������
Sheets = Excel.ActiveWorkBook.Sheets;
%���ص�һ�������
sheet1 = get(Sheets, 'Item', 1);
%�����һ�����
invoke(sheet1, 'Activate');
%�����ǰ����������ͼ�δ��ڣ�ͨ��ѭ����ͼ��ȫ��ɾ��
Shapes=Excel.ActiveSheet.Shapes;
if Shapes.Count~=0;
for i=1:Shapes.Count;
Shapes.Item(1).Delete;
end;
end;

%��ͼ�θ��Ƶ�ճ����
hgexport(fh, '-clipboard');
%��ͼ��ճ������ǰ����A5:B5����
Excel.ActiveSheet.Range('H5:I5').Select;
Excel.ActiveSheet.Paste;
%ɾ��ͼ�ξ��
% 
Workbook.Save();%% �����ļ�
% Workbook.SaveAs(filespec_user);%% �����ļ�
delete(Workbook);

%% �������ɵ�����ͼ  ��ɾ����ǰ���ڵ�ͼƬ
pngFile = strcat(imname,'_tupian.png'); %%��ɴ�excle�ļ�����podu�ļ���
figFile = strcat(imname,'_tupian.fig'); %%��ɴ�excle�ļ�����podu�ļ���
% if exist('tupian.png')   
%     delete('tupian.png');
% end
% if exist('tupian.fig')   
%     delete('tupian.fig');
% end
saveas(gcf,pngFile);
saveas(gcf,figFile);
delete(fh);
%% ���ݴ�����ϣ������ʾ��Ϣ
disp('���ݴ�����ϣ���鿴��ǰ�ļ����µ�');
disp(poduFile);
filePath    %% �ļ�����λ��
cd ..       %%�˳�outputĿ¼

end

