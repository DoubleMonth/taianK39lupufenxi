%% ���ܣ�ʹ��Զ�̼���ռ������ݼ����¶�
%% �����ļ�����Ҫ������ļ�����ͬһ���ļ����£�����Զ�̼��ϵͳ���������ݴ򿪺����Ϊ.xlsx��ʽ��ɾ��������excel�ļ���
%% --��Ϊ����������Ϊ�ı���ʽ��excel�ļ���MATLAB����ʱ�����
%% �޸�filename���ļ���Ϊ���ݵ��ļ���
%% �������Ҫ�������ݣ���ע�͵��˲��㷨���֣������Ҫ���˵����ݲ���0.2�������˲��㷨���޸���Ӧ����
%% ����ɹ����ڵ�ǰ�ļ��������podu.xls���ļ���
%% ע�⣺���벻Ҫ�����ļ����������ĵ�·����ʹ�ã��������ִ��󣻳�������ǰ��ر��Ѿ��򿪵�podu.xls�ļ���

%% ��   ����(20190516)V1.0
%% ��   �ߣ�
%% �޸�ʱ�䣺2019-5-16
clear           %% ��������ռ��е����б�����
clc
path = pwd;
dirOutput = dir(fullfile(path,'*.xlsx'));
fileName = {dirOutput.name};
%% ��ʾ��Ҫ������ļ�
fprintf('���ҵ�%d���ļ���Ҫ����ֱ��ǣ�',length(fileName));
fprintf('\n');
for i=1:length(fileName)
    fprintf(strcat(fileName{1,i},'\n'));
end
for i=1:length(fileName)
    disp(sprintf('���ڴ����%d/%d���ļ�',i,length(fileName)));  %���ڴ����x/X���ļ�
    filename=fileName{1,i};%%�������ݱ�������
    f_calculateGradient(filename);
end
system('taskkill /F /IM EXCEL.EXE');% ����excel����