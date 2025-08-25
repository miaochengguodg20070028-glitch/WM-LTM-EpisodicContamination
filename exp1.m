clear all;
clc;

% 指定文件夹路径
folderPath = 'F:\mcg_DATA\LYL\result-lyl\result-lyl\LTM-18-xlsx\18';

% 结果保存文件路径
resultsFilename = fullfile(folderPath, 'fit_results.xlsx');

% 如果结果文件存在，则删除它
if exist(resultsFilename, 'file') == 2
    delete(resultsFilename);
end

% 获取文件夹下所有 Excel 文件
excelFiles = dir(fullfile(folderPath, '*.xlsx'));

% 初始化结果表
allResults = table('Size', [0, 4], 'VariableTypes', {'string', 'double', 'double', 'double'}, ...
    'VariableNames', {'Filename', 'MaxPosterior', 'LowerCredible', 'UpperCredible'});

% 遍历每个 Excel 文件
for i = 1:length(excelFiles)
    try
        % 构建完整文件路径
        filename = fullfile(folderPath, excelFiles(i).name);
        
        % 读取数据（不将第一行作为表头）
        dataTable = readtable(filename, 'ReadVariableNames', false);

        % 检查第 12 列是否存在且不为空
        if size(dataTable, 2) < 12 || isempty(dataTable{:, 12})
            fprintf('文件 %s 没有第 11 列数据或数据为空，跳过该文件。\n', excelFiles(i).name);
            continue;
        end

        % 提取第 12 列数据
        data = dataTable{:, 12};
        data = data(~isnan(data)); % 去除 NaN
        data = data(:); % 确保是列向量

        % 检查数据是否有效
        if isempty(data)
            fprintf('文件 %s 的第 12 列数据无效，跳过该文件。\n', excelFiles(i).name);
            continue;
        end

        % 检查数据范围
        if any(data < -180) || any(data > 180)
            fprintf('文件 %s 的第 12 列数据超出范围 [-180, 180]，跳过该文件。\n', excelFiles(i).name);
            continue;
        end

        % 打印数据信息
        fprintf('文件 %s 的第 12 列数据：\n', excelFiles(i).name);
        disp(data);
        fprintf('最小值：%.2f，最大值：%.2f\n', min(data), max(data));

        % 将数据打包进一个结构体
        dataStruct = struct();
        dataStruct.errors = data; % 第 12 列数据

        % 打印模型输入
        disp('dataStruct.errors:'); 
        disp(dataStruct.errors);

        % 创建模型结构体
        model = StandardMixtureModel();

        % 打印模型参数
        disp('模型参数：');
        disp(model);

        % 调用 MemFit 进行模型拟合
        fit = MemFit(dataStruct, model);
        
        % 检查拟合结果是否包含所需字段
        if isfield(fit, 'maxPosterior') && isfield(fit, 'lowerCredible') && isfield(fit, 'upperCredible')
            % 提取 maxPosterior、lowerCredible 和 upperCredible
            maxPosterior = fit.maxPosterior;
            lowerCredible = fit.lowerCredible;
            upperCredible = fit.upperCredible;
        else
            fprintf('文件 %s 的拟合结果不完整，跳过该文件。\n', excelFiles(i).name);
            continue;
        end

        % 显示拟合结果
        disp(fit);

        % 构建保存结果的数据表
        resultsTable = table({excelFiles(i).name}, maxPosterior, lowerCredible, upperCredible, ...
            'VariableNames', {'Filename', 'MaxPosterior', 'LowerCredible', 'UpperCredible'});
        
        % 将当前结果追加到总结果表中
        allResults = [allResults; resultsTable];
    catch ME
        % 捕获错误并显示
        fprintf('Error processing file: %s\n', excelFiles(i).name);
        fprintf('Error message: %s\n', ME.message);
    end
end

% 将所有结果一次性写入 Excel 文件
writetable(allResults, resultsFilename, 'Sheet', 1);