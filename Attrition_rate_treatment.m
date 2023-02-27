%    工作簿: C:\Users\gyj92\Desktop\2021\C\附件2 近5年8家转运商的相关数据.xlsx
% 由 MATLAB于09-10 07:55:01 生成
%% 设置选项
opts = spreadsheetImportOptions("NumVariables", 240);
% 指定工作表和范围
opts.Sheet = "运输损耗率（%）";
opts.DataRange = "B2:IG9";
% 指定列与数据类型
opts.VariableNames = ["W001", "W002", "W003", "W004", "W005", "W006", "W007", "W008", "W009", "W010", "W011", "W012", "W013", "W014", "W015", "W016", "W017", "W018", "W019", "W020", "W021", "W022", "W023", "W024", "W025", "W026", "W027", "W028", "W029", "W030", "W031", "W032", "W033", "W034", "W035", "W036", "W037", "W038", "W039", "W040", "W041", "W042", "W043", "W044", "W045", "W046", "W047", "W048", "W049", "W050", "W051", "W052", "W053", "W054", "W055", "W056", "W057", "W058", "W059", "W060", "W061", "W062", "W063", "W064", "W065", "W066", "W067", "W068", "W069", "W070", "W071", "W072", "W073", "W074", "W075", "W076", "W077", "W078", "W079", "W080", "W081", "W082", "W083", "W084", "W085", "W086", "W087", "W088", "W089", "W090", "W091", "W092", "W093", "W094", "W095", "W096", "W097", "W098", "W099", "W100", "W101", "W102", "W103", "W104", "W105", "W106", "W107", "W108", "W109", "W110", "W111", "W112", "W113", "W114", "W115", "W116", "W117", "W118", "W119", "W120", "W121", "W122", "W123", "W124", "W125", "W126", "W127", "W128", "W129", "W130", "W131", "W132", "W133", "W134", "W135", "W136", "W137", "W138", "W139", "W140", "W141", "W142", "W143", "W144", "W145", "W146", "W147", "W148", "W149", "W150", "W151", "W152", "W153", "W154", "W155", "W156", "W157", "W158", "W159", "W160", "W161", "W162", "W163", "W164", "W165", "W166", "W167", "W168", "W169", "W170", "W171", "W172", "W173", "W174", "W175", "W176", "W177", "W178", "W179", "W180", "W181", "W182", "W183", "W184", "W185", "W186", "W187", "W188", "W189", "W190", "W191", "W192", "W193", "W194", "W195", "W196", "W197", "W198", "W199", "W200", "W201", "W202", "W203", "W204", "W205", "W206", "W207", "W208", "W209", "W210", "W211", "W212", "W213", "W214", "W215", "W216", "W217", "W218", "W219", "W220", "W221", "W222", "W223", "W224", "W225", "W226", "W227", "W228", "W229", "W230", "W231", "W232", "W233", "W234", "W235", "W236", "W237", "W238", "W239", "W240"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];
% 导入数据
yuan = readtable("C:\Users\gyj92\Desktop\2021\C\附件2 近5年8家转运商的相关数据.xlsx", opts, "UseExcel", false);
%% 转换为矩阵形式
yuan = table2array(yuan); 

%% 清除变量
clear  opts
%% 处理
lin = zeros(8,240);
index = zeros(8,1);
for i = 1 :8
    for j = 1:240
        if yuan(i,j) ~= 0
            index(i,1) = index(i,1)+1;
            lin(i,index(i,1)) = yuan(i,j);

        end
    end
    plot(lin(i,1:index(i,1)));
    hold on;
end
legend("T"+string(1),"T"+string(2),"T"+string(3),"T"+string(4),"T"+string(5),"T"+string(6),"T"+string(7),"T"+string(8));
figure1 = figure;
for i =1:8
    % 创建 随图，将八个供应商以2*4分布
    subplot1 = subplot(2,4,i,'Parent',figure1);
    hold(subplot1,'on');
    plot(lin(i,1:index(i,1)));
    plot(lin(i,1:index(i,1)),'Parent',subplot1,'LineWidth',1,...
    'Color',[0 0.450980392156863 0.741176470588235]);
    % 创建 y轴名称
    ylabel({'损耗率（%）'});
    % 创建 x轴名称
    xlabel({'周序'});
    % 创建每个图标题
    title({'转运商T'+string(i)});
end
%% 处理五年数据
yuanm = zeros(8,48);
for i = 1:8
    for j = 1:48
        lin = [yuan(i,j),yuan(i,48+j),yuan(i,96+j),yuan(i,144+j),yuan(i,192+j)];
        yuanm(i,j) = sum(lin)/(5-sum(ismember(0,lin)));
    end
end
clear lin;
yuanjun = zeros(1,48);
for i = 1:48
    yuanjun(i) = sum(yuanm(:,i))/(8-sum(ismember(0,yuanm(:,i))));
end
xlswrite("yuan.xlsx",yuanjun);