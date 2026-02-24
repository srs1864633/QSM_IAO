%% LAS File Processing and QSM Model Optimization Pipeline
% This script processes multiple LAS point cloud files, generates QSM (Quantitative Structure Model) 
% models with randomized parameters, selects optimal models, and saves results to Excel and MAT files.

%% Define file paths (modify according to your environment)
inputFolderPath = 'your_input_folder_path';  % Folder containing LAS files
outputFolderPath = 'your_output_folder_path'; % Folder for saving results

%% Retrieve all LAS files in the folder
lasFiles = dir(fullfile(inputFolderPath, '*.las'));

% Check if any .las files are found
if isempty(lasFiles)
    error('No LAS files found in the specified directory.');
end

% Display total number of files to process
fprintf('Total %d LAS files to process.\n', length(lasFiles));

%% Ensure output directory exists
if ~exist(outputFolderPath, 'dir')
    mkdir(outputFolderPath);
end

%% Process each LAS file sequentially
for fileIndex = 1:length(lasFiles)
    % Display current processing status
    fprintf('Processing file %d of %d: %s\n', fileIndex, length(lasFiles), lasFiles(fileIndex).name);
    
    % Initialize structure for point cloud data
    ptCloudData = struct();
    
    % Get full path of current LAS file
    lasFilePath = fullfile(lasFiles(fileIndex).folder, lasFiles(fileIndex).name);
    lasReader = lasFileReader(lasFilePath);
    
    % Read point cloud data
    ptCloud = readPointCloud(lasReader);
    
    % Center point cloud by subtracting mean position
    ptCloudLocations = ptCloud.Location;
    ptCloudLocations = ptCloudLocations - mean(ptCloudLocations);
    
    % Extract tree name from filename (without extension)
    [~, treeName, ~] = fileparts(lasFiles(fileIndex).name);
    
    % Save processed point cloud to structure
    ptCloudData.(treeName) = ptCloudLocations;
    
    % Save structure to MAT file
    save('trees', '-struct', 'ptCloudData');
    
    %% Define randomized input parameters for QSM generation
    % Clear previous input structure
    inputs = struct();
    
    % Patch size for first uniform-size cover (randomized between 0.05 and 0.5)
    inputs.PatchDiam1 = (0.5 - 0.05) * rand(1, 2) + 0.05;
    
    % Minimum patch size for second cover (randomized between 0.02 and 0.2)
    inputs.PatchDiam2Min = (0.2 - 0.02) * rand(1, 4) + 0.02;
    
    % Maximum patch size at stem base in second cover (randomized between 0.1 and 0.6)
    inputs.PatchDiam2Max = (0.6 - 0.1) * rand(1, 3) + 0.1;
    
    % Additional patch generation parameters
    inputs.BallRad1 = inputs.PatchDiam1 + 0.015;
    inputs.BallRad2 = inputs.PatchDiam2Max + 0.01;
    
    %% Fixed parameter settings
    inputs.nmin1 = 3;
    inputs.nmin2 = 1;
    inputs.OnlyTree = 1;
    inputs.Tria = 0;
    inputs.Dist = 1;
    inputs.MinCylRad = 0.0025;
    inputs.ParentCor = 1;
    inputs.TaperCor = 1;
    inputs.GrowthVolCor = 0;
    inputs.GrowthVolFac = 1.5;
    
    %% Filter parameter settings
    inputs.filter.k = 10;
    inputs.filter.radius = 0.00;
    inputs.filter.nsigma = 1.5;
    inputs.filter.PatchDiam1 = 0.05;
    inputs.filter.BallRad1 = 0.075;
    inputs.filter.ncomp = 2;
    inputs.filter.EdgeLength = 0.004;
    inputs.filter.plot = 1;
    
    %% Additional input settings
    inputs.name = 'tree';
    inputs.tree = 1;
    inputs.model = 1;
    inputs.savemat = 1;
    inputs.savetxt = 1;
    inputs.plot = 2;
    inputs.disp = 2;
    
    %% Generate QSM models for current tree
    QSMs = make_models('trees', treeName, 1, inputs);
    
    % Select optimal model
    [TreeData, OptModels, OptInputs, OptQSM] = select_optimum(QSMs);
    
    %% Generate results table
    columnNames = {'Name', 'NameLetters', 'NameNumbers', 'PatchDiam1', 'PatchDiam2Min', 'PatchDiam2Max', ...
        'BallRad1', 'BallRad2', 'TotalVolume', 'TrunkVolume', 'BranchVolume', ...
        'TreeHeight', 'TrunkLength', 'BranchLength', 'TotalLength', 'NumberBranches', 'MaxBranchOrder', ...
        'TrunkArea', 'BranchArea', 'TotalArea', 'DBHqsm', 'DBHcyl', 'CrownDiamAve', 'CrownDiamMax', ...
        'CrownAreaConv', 'CrownAreaAlpha', 'CrownBaseHeight', 'CrownLength', 'CrownRatio', ...
        'CrownVolumeConv', 'CrownVolumeAlpha', 'Runtime'};
    
    resultsTable = table();
    
    % Populate table with data from all QSM models
    for modelIndex = 1:numel(QSMs)
        rundata = QSMs(modelIndex).rundata.inputs;
        treedata = QSMs(modelIndex).treedata;
        runtime = QSMs(modelIndex).rundata.time;
        
        % Extract letters and numbers from name
        tokens = regexp(rundata.name, '^([a-zA-Z]+)(\d+)$', 'tokens');
        if ~isempty(tokens)
            nameLetters = tokens{1}{1};
            nameNumbers = str2double(tokens{1}{2});
        else
            nameLetters = '';
            nameNumbers = NaN;
        end
        
        % Create new row for table
        newRow = {rundata.name, nameLetters, nameNumbers, rundata.PatchDiam1, rundata.PatchDiam2Min, rundata.PatchDiam2Max, ...
            rundata.BallRad1, rundata.BallRad2, treedata.TotalVolume, treedata.TrunkVolume, treedata.BranchVolume, ...
            treedata.TreeHeight, treedata.TrunkLength, treedata.BranchLength, ...
            treedata.TotalLength, treedata.NumberBranches, treedata.MaxBranchOrder, ...
            treedata.TrunkArea, treedata.BranchArea, treedata.TotalArea, treedata.DBHqsm, ...
            treedata.DBHcyl, treedata.CrownDiamAve, treedata.CrownDiamMax, ...
            treedata.CrownAreaConv, treedata.CrownAreaAlpha, treedata.CrownBaseHeight, ...
            treedata.CrownLength, treedata.CrownRatio, treedata.CrownVolumeConv, ...
            treedata.CrownVolumeAlpha, runtime};
        
        resultsTable = [resultsTable; newRow];
    end
    
    % Set table column names
    resultsTable.Properties.VariableNames = columnNames;
    
    % Generate filename for output
    fileName = sprintf('QSM_%s_%f-%f-%f', rundata.name, rundata.PatchDiam1, rundata.PatchDiam2Min, rundata.PatchDiam2Max);
    fileName = strrep(fileName, '.', '-');
    fileName = strrep(fileName, ' ', '_');
    
    % Save Excel file
    excelFilePath = fullfile(outputFolderPath, [fileName '.xlsx']);
    writetable(resultsTable, excelFilePath, 'Sheet', 'QSMs');
    
    % Save MAT file
    matFilePath = fullfile(outputFolderPath, [fileName '.mat']);
    save(matFilePath, 'QSMs');
    
    %% Generate OptQSM results table
    optQSMTable = table();
    
    for optIndex = 1:numel(OptQSM)
        rundata = OptQSM(optIndex).rundata.inputs;
        treedata = OptQSM(optIndex).treedata;
        runtime = OptQSM(optIndex).rundata.time;
        
        tokens = regexp(rundata.name, '^([a-zA-Z]+)(\d+)$', 'tokens');
        if ~isempty(tokens)
            nameLetters = tokens{1}{1};
            nameNumbers = str2double(tokens{1}{2});
        else
            nameLetters = '';
            nameNumbers = NaN;
        end
        
        newRow = {rundata.name, nameLetters, nameNumbers, rundata.PatchDiam1, rundata.PatchDiam2Min, rundata.PatchDiam2Max, ...
            rundata.BallRad1, rundata.BallRad2, treedata.TotalVolume, treedata.TrunkVolume, treedata.BranchVolume, ...
            treedata.TreeHeight, treedata.TrunkLength, treedata.BranchLength, ...
            treedata.TotalLength, treedata.NumberBranches, treedata.MaxBranchOrder, ...
            treedata.TrunkArea, treedata.BranchArea, treedata.TotalArea, treedata.DBHqsm, ...
            treedata.DBHcyl, treedata.CrownDiamAve, treedata.CrownDiamMax, ...
            treedata.CrownAreaConv, treedata.CrownAreaAlpha, treedata.CrownBaseHeight, ...
            treedata.CrownLength, treedata.CrownRatio, treedata.CrownVolumeConv, ...
            treedata.CrownVolumeAlpha, runtime};
        
        optQSMTable = [optQSMTable; newRow];
    end
    
    % Set OptQSM table column names
    optQSMTable.Properties.VariableNames = columnNames;
    
    % Save OptQSM table to new sheet
    writetable(optQSMTable, excelFilePath, 'Sheet', 'OptQSM');
end

%% Merge all generated Excel files
% Get all Excel files in output folder
excelFiles = dir(fullfile(outputFolderPath, '*.xlsx'));

% Initialize merged tables
mergedResultsTable_QSM = table();
mergedResultsTable_OptQSM = table();

% Merge data from all Excel files
for mergeIndex = 1:length(excelFiles)
    currentFile = fullfile(excelFiles(mergeIndex).folder, excelFiles(mergeIndex).name);
    
    % Read QSMs sheet
    QSMData = readtable(currentFile, 'Sheet', 'QSMs');
    mergedResultsTable_QSM = [mergedResultsTable_QSM; QSMData];
    
    % Read OptQSM sheet
    OptQSMData = readtable(currentFile, 'Sheet', 'OptQSM');
    mergedResultsTable_OptQSM = [mergedResultsTable_OptQSM; OptQSMData];
end

% Save merged tables to new Excel file
mergedFileName = fullfile(outputFolderPath, 'Merged_QSM_OptQSM.xlsx');
writetable(mergedResultsTable_QSM, mergedFileName, 'Sheet', 'QSMs');
writetable(mergedResultsTable_OptQSM, mergedFileName, 'Sheet', 'OptQSM');

fprintf('All files merged and saved to "Merged_QSM_OptQSM.xlsx"\n');