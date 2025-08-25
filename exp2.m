clear all
clc

% 指定文件夹路径
folderPath = 'D:\资料\博士课程\华师\WSY\LYL\data\result-lyl\all-wl\all-wl-repeation\1';

% 结果保存文件路径
resultsFilename = fullfile(folderPath, 'fit_results.xlsx');

% 如果结果文件存在，则删除它
if exist(resultsFilename, 'file') == 2
    delete(resultsFilename);
end

% 获取文件夹下所有 Excel 文件
excelFiles = dir(fullfile(folderPath, '*.xlsx'));

% 初始化行计数器
rowCounter = 1;

% 初始化结果表
allResults = table('Size', [0, 4], 'VariableTypes', {'string', 'double', 'double', 'double'}, ...
    'VariableNames', {'Filename', 'MaxPosterior', 'LowerCredible', 'UpperCredible'});

% 遍历每个 Excel 文件
for i = 1:length(excelFiles)
    % 构建完整文件路径
    filename = fullfile(folderPath, excelFiles(i).name);
    
    % 读取数据
    data = readtable(filename);

    % 构建数据结构
    dataStruct = struct();
    dataStruct.errors = data{:, 55}; % 第一列是 errors
    dataStruct.distractors = data{:, 56}; % 第二列是 distractors

    % 创建模型
    model = SwapModel;
    
    fit = MemFit(dataStruct, model);

    % 显示拟合结果
    disp(fit);

    % 提取 maxPosterior、lowerCredible 和 upperCredible
    maxPosterior = fit.maxPosterior;
    lowerCredible = fit.lowerCredible;
    upperCredible = fit.upperCredible;

    % 构建保存结果的数据表
    resultsTable = table({excelFiles(i).name}, maxPosterior, lowerCredible, upperCredible, ...
        'VariableNames', {'Filename', 'MaxPosterior', 'LowerCredible', 'UpperCredible'});
    
    % 将当前结果追加到总结果表中
    allResults = [allResults; resultsTable];
    
% 将所有结果一次性写入 Excel 文件
writetable(allResults, resultsFilename, 'Sheet', 1);
end