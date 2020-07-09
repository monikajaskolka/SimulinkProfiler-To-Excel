function profilerToExcel(profileData, filename)
% PROFILERTOEXCEL Convert Simulink Profiler data into an Excel spreadsheet.
%   Currently, the Simulink Profiler only supports exporting data to a MAT file.
%   You can load the MAT data into your workspace, and then use this script to
%   convert it into an Excel file.
%
%   For more information about the Simulink Profiler, see:
%   https://www.mathworks.com/help/simulink/slref/simulinkprofiler.html
%
%   Inputs:
%       profileData     Simulink.profiler.Data object.
%       filename        Desired Excel filename.
%
%   Outputs:
%       N/A
%
%   Side Effects:
%       Excel file created in the present working directory.
%
%   Example:
%       profilerToExcel(profileData, 'text.xlsx');
%
%   Author:
%       Monika Jaskolka

    % Add extension if not provided
    [~, ~, ext] = fileparts(filename);
    if isempty(ext)
        filename = [filename '.xlsx'];
    end

    header = {'Name', 'Total Time (s)', 'Self Time (s)', 'Number of Calls'};
    [name, totalTime, selfTime, numberOfCalls] = profilerDataToCells(profileData);
    data = horzcat(name', num2cell(totalTime)', num2cell(selfTime)', num2cell(numberOfCalls)');

    exceldata = vertcat(header, data);
    xlswrite(filename, exceldata);
end

function [name, totalTime, selfTime, numberOfCalls] = profilerDataToCells(profilerData)
% PROFILERDATATOCELLS Extract the profilng data as cell arrays.
   rootnode = profilerData;
   rootnode = rootnode.rootUINode;
   [name, totalTime, selfTime, numberOfCalls] = recurseProfilerData(rootnode);
end

function [name, totalTime, selfTime, numberOfCalls] = recurseProfilerData(node)
% RECURSEPROFILERDATA Find the data of profile nodes.

    %% Node path
    nodepath = node.path;

    % If the block path has nested Models, then there are multiple paths.
    % Combine the paths into one, w.r.t. the top-most model.
    p_fixed = '';
    if numel(nodepath) > 1
        p_fixed = char(nodepath(1));
        for i = 2:numel(nodepath)
            p = char(nodepath(i));
            slash = strfind(p, '/');
            if isempty(slash)
                endOfModelName = length(p);
            else
                endOfModelName = slash - 1;
            end
            mdlname = p(1:endOfModelName);
            p = strrep(p, mdlname, [' (' mdlname ')']);
            p_fixed = [p_fixed, p];
        end

        name = p_fixed;
    else
        name = char(nodepath);
    end
    name = {name}; % Convert to a cell so we can concatenate as an array

    %% Node time
    totalTime = node.totalTime;
    selfTime = node.selfTime;
    numberOfCalls = node.numberOfCalls;

    %% Base case
    if ~hasChildren(node)
        return;
    end

    %% Recursion
    for j = 1:length(node.children)
        [n, t, s, c] = recurseProfilerData(node.children(j));

        % Append
        name = horzcat(name, n);
        totalTime = [totalTime, t];
        selfTime = [selfTime, s];
        numberOfCalls = [numberOfCalls, c];
    end
end

function c = hasChildren(node)
% HASCHILDREN Determine if the node has children.
%
%   Inputs:
%       node    UINode object.
%
%   Outputs:
%       c       Whether the node has children(1) or not(0).

    try
        c = ~isempty(node.children);
    catch
        c = false;
    end
end