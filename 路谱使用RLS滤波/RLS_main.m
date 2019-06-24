clear
clc
path = pwd;
dirOutput = dir(fullfile(path,'*.xls'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%输入数据表格的名称
[excelData,str] = xlsread(filename,1);               %读取原始数据表中的数据：str为数据表中的字符，data为数据表中的数据
[excelRow,excelColumn] = size(excelData); 

ref_noise=randn(1,excelRow);
mixed = excelData(:,4);

mu=0.05;M=2;espon=1e-4;
delta = 1e-7;
lambda = 1;
[en,w]=rls(lambda,M,ref_noise,mixed,delta);
% [enn,ww]=rls(lambda,M,ref_noise,en,delta);
subplot(4,1,1);
title('噪声');
plot(ref_noise,'r-');
subplot(4,1,2);
title('原始坡度');
plot(mixed,'r-');
subplot(4,1,3);
title('滤波后坡度');
plot(en,'r-');

i = find('.'==filename);
imname = filename(1:i-1); %% imname为不带后缀文件名称 
outFile = strcat(imname,'_output');
if exist(outFile)   %% 如果存在output文件夹，先删除
     rmdir (outFile,'s');
end
mkdir(outFile);%% 创建一个Output文件夹
cd(fullfile(path,outFile));       %%进入output目录
poduFile = strcat(imname,'_RLS.xls'); %%组成带excle文件名的podu文件名
colname={'序号','时间','仪表车速计算坡度','仪表车速计算坡度滤波后','累计里程计算坡度','累计里程计算坡度滤波后','GPS车速计算坡度','GPS车速计算坡度滤波后','GPS里程计算坡度','GPS车速计算坡度滤波后'};    %%增加每一列的数据名称
warning off MATLAB:xlswrite:AddSheet;   %%防止出现warning警告 
xlswrite(poduFile, colname, 'sheet1','A1');
xlswrite(poduFile, excelData(:,1), 'sheet1','A2');
xlswrite(poduFile, str(2:length(str)-1,2), 'sheet1','B2');
xlswrite(poduFile, excelData(:,3), 'sheet1','C2');
xlswrite(poduFile, en', 'sheet1','D2');


