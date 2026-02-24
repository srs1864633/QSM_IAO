%% Main Steps of the IAO Method
% This script reads data from two Excel files, calculates relative RMSE values,
% and selects the parameter value with minimal average rRMSE for each unique Name.

%% Define file paths (modify these according to your environment)
mainDataPath = 'your_main_data_file.xlsx';        % Main data file path
referenceDataPath = 'your_reference_data.xlsx';   % Reference data file path
mainSheetName = 'Sheet1';                         % Main data sheet name
outputSheetName = 'IAO';                          % Output sheet name

%% Read data from Excel files
% Read main data from specified sheet
mainData = readtable(mainDataPath, 'Sheet', mainSheetName);

% Read reference data (using default sheet)
referenceData = readtable(referenceDataPath);

%% Extract unique names from main data
uniqueNames = unique(mainData.Name);

%% Initialize result table
resultTable = table();

%% Process each unique name
for nameIdx = 1:length(uniqueNames)
    currentName = uniqueNames{nameIdx};
    
    % Extract rows with current name from main data
    nameMaskMain = strcmp(mainData.Name, currentName);
    currentMainData = mainData(nameMaskMain, :);
    
    % Find matching reference data for current name
    nameMaskRef = strcmp(referenceData.Name, currentName);
    
    if sum(nameMaskRef) > 0
        % Extract reference data (assuming one matching row)
        refRow = referenceData(nameMaskRef, :);
        
        % Calculate rRMSE for height using reference lidarh
        rRMSE_height = sqrt(mean((refRow.lidarh(1) - currentMainData.qsmh).^2)) / refRow.lidarh(1);
        
        % Calculate rRMSE for DBH using reference lidardbh
        rRMSE_dbh = sqrt(mean((refRow.lidardbh(1) - currentMainData.qsmdbh).^2)) / refRow.lidardbh(1);
        
        % Calculate average rRMSE
        averageRMSE = (rRMSE_height + rRMSE_dbh) / 2;
        
        % Add average rRMSE as temporary column
        currentMainData.rRMSE_mean = repmat(averageRMSE, height(currentMainData), 1);
        
        % Find row with minimal average rRMSE
        [~, minIndex] = min(currentMainData.rRMSE_mean);
        selectedRow = currentMainData(minIndex, :);
        
        % Append selected row to result table
        resultTable = [resultTable; selectedRow];
    else
        % Warning if no matching reference data found
        warning('Name %s not found in reference data', currentName);
    end
end

%% Finalize and save results
if height(resultTable) > 0
    % Remove temporary rRMSE_mean column
    resultTable.rRMSE_mean = [];
    
    % Write results to new sheet in main data file
    writetable(resultTable, mainDataPath, 'Sheet', outputSheetName);
    
    % Display completion message
    fprintf('Processing complete. Results saved to %s sheet.\n', outputSheetName);
else
    disp('No matching data found. Result table is empty.');
end