clear
clc
path = pwd;
dirOutput = dir(fullfile(path,'*.xls'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%输入数据表格的名称
[excelData,str] = xlsread(filename,1);               %读取原始数据表中的数据：str为数据表中的字符，data为数据表中的数据
[excelRow,excelColumn] = size(excelData); 

ref_noise=randn(1,excelRow);
mixed = excelData(:,3);

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
