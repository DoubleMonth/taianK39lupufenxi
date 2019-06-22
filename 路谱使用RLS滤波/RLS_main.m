clear
clc
path = pwd;
dirOutput = dir(fullfile(path,'*.xls'));
fileName = {dirOutput.name};
filename=fileName{1,1};%%�������ݱ�������
[excelData,str] = xlsread(filename,1);               %��ȡԭʼ���ݱ��е����ݣ�strΪ���ݱ��е��ַ���dataΪ���ݱ��е�����
[excelRow,excelColumn] = size(excelData); 

ref_noise=randn(1,excelRow);
mixed = excelData(:,3);

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
