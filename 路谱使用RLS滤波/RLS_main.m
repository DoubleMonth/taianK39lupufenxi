clear
clc
path = pwd;
dirOutput = dir(fullfile(path,'*.xls'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%�������ݱ�������
[excelData,str] = xlsread(filename,1);               %��ȡԭʼ���ݱ��е����ݣ�strΪ���ݱ��е��ַ���dataΪ���ݱ��е�����
[excelRow,excelColumn] = size(excelData); 

ref_noise=randn(1,excelRow);
mixed = excelData(:,4);

mu=0.05;M=2;espon=1e-4;
delta = 1e-7;
lambda = 1;
[en,w]=rls(lambda,M,ref_noise,mixed,delta);
% [enn,ww]=rls(lambda,M,ref_noise,en,delta);
subplot(4,1,1);
title('����');
plot(ref_noise,'r-');
subplot(4,1,2);
title('ԭʼ�¶�');
plot(mixed,'r-');
subplot(4,1,3);
title('�˲����¶�');
plot(en,'r-');

i = find('.'==filename);
imname = filename(1:i-1); %% imnameΪ������׺�ļ����� 
outFile = strcat(imname,'_output');
if exist(outFile)   %% �������output�ļ��У���ɾ��
     rmdir (outFile,'s');
end
mkdir(outFile);%% ����һ��Output�ļ���
cd(fullfile(path,outFile));       %%����outputĿ¼
poduFile = strcat(imname,'_RLS.xls'); %%��ɴ�excle�ļ�����podu�ļ���
colname={'���','ʱ��','�Ǳ��ټ����¶�','�Ǳ��ټ����¶��˲���','�ۼ���̼����¶�','�ۼ���̼����¶��˲���','GPS���ټ����¶�','GPS���ټ����¶��˲���','GPS��̼����¶�','GPS���ټ����¶��˲���'};    %%����ÿһ�е���������
warning off MATLAB:xlswrite:AddSheet;   %%��ֹ����warning���� 
xlswrite(poduFile, colname, 'sheet1','A1');
xlswrite(poduFile, excelData(:,1), 'sheet1','A2');
xlswrite(poduFile, str(2:length(str)-1,2), 'sheet1','B2');
xlswrite(poduFile, excelData(:,3), 'sheet1','C2');
xlswrite(poduFile, en', 'sheet1','D2');


