function cdl = convertNcml(ncmlfile)

    % convert a ncml file in various forms

    pref.Str2Num = 'never';
    d = xml_read(ncmlfile, pref);
    cdl = '';
    tables = struct([]);
    tab = 0;

    %%%%%%%%%%
    % Convert to CDL
    
    % Header
    cdl = [cdl sprintf('netcdf %s {\n', d.ATTRIBUTE.title)];
    t = struct('groupName', 'Toplevel', 'groupText', {{'Description', 'Obligation', 'Comment'}});

    % global attributes
    tab = tab + 1;
    [attrs, tableT] = formatAttributes('', d.attribute, tab);
    cdl = [cdl attrs];
    t.groupText = [t.groupText; {'Group attributes' '' ''}; tableT];
    tab = tab - 1;
    
    % global dimensions
    tab = tab + 1;
    [cdlText, tableCells] = formatDimensions(d.dimension, tab);
    cdl = [cdl cdlText];
    t.groupText = [t.groupText; tableCells];
    tab = tab - 1;

    tables = [tables; t];
    
    % groups
    d.group = sortStruct(d.group, {'ATTRIBUTE' 'name'});
    for group_i = 1:length(d.group) 
        [groupCdl, tableT] = formatGroup(d.group(group_i), tab);
        cdl = [cdl groupCdl]; %#ok<*AGROW>
        tables = [tables; tableT];
    end
    %ty
    cdl = [cdl sprintf('}\n')];

    % Write out the CDL file
    cdlfile = [ncmlfile(1:end-5) '.cdl'];
    fid = fopen(cdlfile, 'w');
    fwrite(fid, cdl);
    fclose(fid);
    disp(['Wrote ' cdlfile])
    
    % write out tables into a Word document
    wordfile=[ncmlfile(1:end-5) '-tables-only.docx'];
    wordfile = GetFullPath(wordfile);
    delete(wordfile); % remove existing file
	[actXWord, wordHandle] = StartWord(wordfile);
    
    % Table heading (convert to caption later on..)
    style='Heading 1';
    text='SONAR-netCDF4 definition';
    WordText(actXWord,text,style,[1,1]); % enter before and after text 
    
    for i = 1:length(tables)
        % Work out which rows to highlight in the table
        highlightRows = [];
        for j = 2:size(tables(i).groupText,1) % don't highlight the first row, ever
            if ~isempty(tables(i).groupText{j,1}) && tables(i).groupText{j,1}(1) ~= ' '
                highlightRows = [highlightRows j];
            end
        end
        captionText = ['Description of the ' tables(i).groupName ' group.'];
        WordCreateTable(actXWord, tables(i).groupText, captionText, highlightRows, 1);
    end
    
    CloseWord(actXWord, wordHandle, wordfile);
    
    disp(['Wrote ' wordfile])
end

function [cdl, t] = formatGroup(group, tab)

    cdl = '';
    t = struct([]);
    if ~isempty(group)
        t = struct('groupName',[], 'groupText', []);
        tableText = {'Description', 'Obligation', 'Comment'};
        tab = tab + 1;
        
        cdl = [cdl sprintf('\n%sgroup: %s {\n', tabs(tab), group.ATTRIBUTE.name)];
        t.groupName = group.ATTRIBUTE.name;
        
        % Attributes
        [cdlText, tableCells] =  formatAttributes('', group.attribute, tab);
        cdl = [cdl cdlText];
        if ~isempty(tableCells)
            tableText = [tableText; {'Group attributes' '' ''}; tableCells];
        end
        
        % Data types
        sEnum = '';
        tEnum = [];
        if isfield(group, 'enumTypedef')
            [sEnum, tEnum] = formatEnumTypes(group.enumTypedef, tab);
        end
        sVar = '';
        tVar = [];
        if isfield(group, 'varTypedef')
            [sVar, tVar] = formatVarTypes(group.varTypedef, tab);
        end
        cdl = [cdl mergeCDLTypes(sEnum, sVar, tab)];
        tableText = [tableText; mergeTableTypes(tEnum, tVar)];

        % Dimensions
        [cdlText, tableCells] = formatDimensions(group.dimension, tab);
        cdl = [cdl cdlText];
        tableText = [tableText; tableCells];
        
        % Variables and Coordinate variables
        [cdlText, tableCells] = formatVariables(group.variable, tab);
        cdl = [cdl cdlText];
        tableText = [tableText; tableCells];
        
        t.groupText = tableText;
        
        % Sub-groups (separate table), but include a table entry
        if isfield(group, 'group') && ~isempty(group.group)
            group.group = sortStruct(group.group, {'ATTRIBUTE' 'name'});
            for group_i = 1:length(group.group)
                [groupCdl, tbl] = formatGroup(group.group(group_i), tab);

                % Add a dummy entry to indicate the sub-group's presence in the
                % parent group.            
                [obligation, comment] = getTableData(group.group.ATTRIBUTE);
                t.groupText = [t.groupText; {'Sub groups' '' ''}; ...
                    {[tabs(1) tbl.groupName] obligation comment}];

                tbl.groupName = [tbl.groupName ' sub'];
                
                cdl = [cdl groupCdl]; %#ok<*AGROW>
                if ~isempty(tbl)
                    t = [t; tbl];
                end
            end
            
            
        end
        
        cdl = [cdl sprintf('%s}\n', tabs(tab))]; % end of group
    end
end

function scdl = mergeCDLTypes(t1, t2, tab)
    % Merges enum and variable type definitions into one 'types:' section. 
    scdl = '';
    
    if ~isempty(t1) || ~isempty(t2)
        tab = tab + 1;
        scdl = [scdl sprintf('\n%stypes:\n', tabs(tab))];
        
        if ~isempty(t1)
            scdl = [scdl t1];
        end
        if ~isempty(t2)
            scdl = [scdl t2];
        end
        scdl = [scdl newline];
    end
end

function t = mergeTableTypes(t1, t2)
    % Merges enum and variable type definitions into one section. 
    t = [];
    
    if ~isempty(t1) || ~isempty(t2)
        t = {'Types', '', ''};
        
        if ~isempty(t1)
            t = [t; t1];
        end
        if ~isempty(t2)
            t = [t; t2];
        end
    end
end

function [cdl, tableText] = formatEnumTypes(types, tab)

    cdl = '';
    tableText = [];
    
    if ~isempty(types)
        tab = tab + 2;

        types = sortStruct(types, {'ATTRIBUTE' 'name'});
        for type_i = 1:length(types)
            t = types(type_i);
            
            [obligation, comment] = getTableData(t.ATTRIBUTE);
            
            base = sprintf('%s enum %s', t.ATTRIBUTE.type, t.ATTRIBUTE.name);
            base = [base ' {'];

            for enum_i = 1:length(t.enum)
                base = [base sprintf('%s = %s, ', t.enum(enum_i).CONTENT, t.enum(enum_i).ATTRIBUTE.key)];
            end
            
            base = base(1:end-2); % trim off trailing comma
            base = [base '}'];
            
            cdl = [cdl sprintf('%s%s;\n', tabs(tab), base)];
            tableText = [tableText; {[tabs(1) base] obligation comment}];
        end
    end
end

function [cdl, tableText] = formatVarTypes(types, tab)

    cdl = '';
    tableText = [];
    
    if ~isempty(types)
        tab = tab + 2;

        types = sortStruct(types, {'ATTRIBUTE' 'name'});
        for type_i = 1:length(types)
            
            t = types(type_i);
            
            [obligation, comment] = getTableData(t.ATTRIBUTE);

            base = sprintf('%s %s', t.ATTRIBUTE.type, t.ATTRIBUTE.name);
            cdl = [cdl sprintf('%s%s;', tabs(tab), base)];
            tableText = [tableText; {[tabs(1) base] obligation comment}];
        end
    end
end

function [cdl, tableText] = formatVariables(variable, tab)

    cdl = '';
    tableText = [];
    
    if ~isempty(variable)
        tab = tab + 1;
        cdl = [cdl sprintf('\n%svariables:\n', tabs(tab))];

        tab = tab + 1;
        
        if length(variable) > 1
            addTrailingLine = true;
        else
            addTrailingLine = false;
        end
        
        % We want to split the variables into two sections, one for
        % coordinate variables and one for normal variables. Both of these
        % sections should be sorted aplhanumerically by variable name.

        % Find out which variables are coordinate ones and which aren't.
        coordinate_var = false(size(variable));
        for variable_i = 1:length(variable)
            var = variable(variable_i).ATTRIBUTE;
            if isfield(var, 'shape') && strcmp(var.name, var.shape)
                coordinate_var(variable_i) = true;
            end
        end

        tableText = [];
        for coord_var_flag = [true, false] % do coordinate vars first
            % select out the type of variables we want
            vars = variable(coordinate_var == coord_var_flag);
            % sort them
            vars = sortStruct(vars, {'ATTRIBUTE' 'name'});
            
            % loop over them and generate our output
            for variable_i = 1:length(vars)
                var = vars(variable_i).ATTRIBUTE;

                if variable_i == 1
                    if coord_var_flag
                        tableText = [tableText; {'Coordinate variables', '', ''}];
                    else
                        tableText = [tableText; {'Variables', '', ''}];
                    end
                end
                
                [obligation, comment] = getTableData(var);
                
                varType = lower(var.type);
                if isfield(var, 'shape')
                    % replace spaces in shape text with commas
                    shape = regexprep(var.shape, '\s+', ', ');
                    base = sprintf('%s %s(%s)', varType, var.name, shape);
                    cdl = [cdl sprintf('%s%s;\n', tabs(tab), base)];
                    tableText = [tableText; {[tabs(1) base] obligation comment}];
                else
                    cdl = [cdl sprintf('%s%s %s;\n', tabs(tab), varType, var.name)];
                    tableText = [tableText; {[tabs(1) varType ' ' var.name] obligation comment}];
                end
                
                [attrs, tableAttr] = formatAttributes(vars(variable_i).ATTRIBUTE.name, ...
                    vars(variable_i).attribute, tab);
                
                cdl = [cdl attrs];
                tableText = [tableText; tableAttr];
                
                if addTrailingLine && variable_i ~= length(vars)
                    tableText = [tableText; {'' '' ''}];
                end
            end
        end
    end
end

function [cdl, tableAttr] = formatAttributes(name, attributes, tab)

    % Outputs formatted attributes for the given variable/group
    cdl = '';
    tableAttr = {};
    tab = tab + 1;

    attributes = sortStruct(attributes, {'ATTRIBUTE' 'name'});
    for attr_i = 1:length(attributes)
        attribute = attributes(attr_i).ATTRIBUTE;
        
        % skip attributes named '_Unsigned'
        if ~strcmp(attribute.name, '_Unsigned')
        
            if isfield(attribute, 'type')
                var_type = [attribute.type ' '];
            else
                var_type = '';
            end
            
            [obligation, comment] = getTableData(attribute);
            
            % If the data type is numeric, treat it appropriately. Take
            % care with space-seperated numeric values. Otherwise treat the
            % value as text to be quoted. 
            if isfield(attribute, 'type') && (strcmp(attribute.type, 'float') || strcmp(attribute.type, 'double'))
                value = regexprep(attribute.value, '\s+', ', ');
            elseif isfield(attribute, 'type') && strcmp(attribute.type(end-1:end), '_t')
                % special case for an enumeration type'd attribute
                value = attribute.value;
            else
                value = ['"' attribute.value '"'];
            end
            base = sprintf(':%s = %s', attribute.name, value);
            cdl = [cdl sprintf('%s%s%s%s;\n', tabs(tab), ...
                var_type, name, base)];
            tableAttr = [tableAttr; {[tabs(2) var_type base] obligation comment}];
        end
    end
end

function [cdl, t] = formatDimensions(dimension, tab)

    cdl = '';
    t = {};
    if ~isempty(dimension)
        tab = tab + 1;
        cdl = [cdl sprintf('\n%sdimensions:\n', tabs(tab))];
        t = {'Dimensions', '', ''};

        dimension = sortStruct(dimension, {'ATTRIBUTE' 'name'});
        for dim_i = 1:length(dimension)
            dim = dimension(dim_i);

            [obligation, comment] = getTableData(dim.ATTRIBUTE);
            
            tab = tab + 1;
            if isfield(dim.ATTRIBUTE, 'isUnlimited') && strcmp(dim.ATTRIBUTE.isUnlimited, 'true')
                base = sprintf('%s = unlimited', dim.ATTRIBUTE.name);
            else
                base = sprintf('%s = %s', dim.ATTRIBUTE.name, dim.ATTRIBUTE.length);
            end
            cdl = [cdl sprintf('%s%s;\n', tabs(tab), base)];
            t = [t; {sprintf('%s%s', tabs(1), base) obligation comment}];

            tab = tab - 1;
        end
    end 
end


function s = tabs(tab)

    s = '';
    for i = 1:tab
        s = [s sprintf('   ')];
    end

end

function [obligation, comment] = getTableData(t)

    obligation = '';
    comment = '';

    if isfield(t, 'obligation')
        obligation = t.obligation;
    end
    if isfield(t, 'comment')
        comment = t.comment;
    end
end

function s = sortStruct(s, key)
    % sort group array alphanumerically based on the structure value given
    % by key(s).
    
    if length(s) > 1
        for i = 1:length(s)
            names{i} = getfield(s, {i}, key{:});
        end
    
        [~, ind] = sort(upper(names));
        s = s(ind);
    end
end


