%    ������: C:\Users\gyj92\Desktop\2021\C\����1 ��5��402�ҹ�Ӧ�̵��������.xlsx
% �� MATLAB ��09-10 09:08:32 ����
%% ����ѡ��
opts = spreadsheetImportOptions("NumVariables", 240);
% ָ����Χ
opts.Sheet = "��ҵ�Ķ�������m?��";
opts.DataRange = "C2:IH403";
% ָ��������������
opts.VariableNames = ["W001", "W002", "W003", "W004", "W005", "W006", "W007", "W008", "W009", "W010", "W011", "W012", "W013", "W014", "W015", "W016", "W017", "W018", "W019", "W020", "W021", "W022", "W023", "W024", "W025", "W026", "W027", "W028", "W029", "W030", "W031", "W032", "W033", "W034", "W035", "W036", "W037", "W038", "W039", "W040", "W041", "W042", "W043", "W044", "W045", "W046", "W047", "W048", "W049", "W050", "W051", "W052", "W053", "W054", "W055", "W056", "W057", "W058", "W059", "W060", "W061", "W062", "W063", "W064", "W065", "W066", "W067", "W068", "W069", "W070", "W071", "W072", "W073", "W074", "W075", "W076", "W077", "W078", "W079", "W080", "W081", "W082", "W083", "W084", "W085", "W086", "W087", "W088", "W089", "W090", "W091", "W092", "W093", "W094", "W095", "W096", "W097", "W098", "W099", "W100", "W101", "W102", "W103", "W104", "W105", "W106", "W107", "W108", "W109", "W110", "W111", "W112", "W113", "W114", "W115", "W116", "W117", "W118", "W119", "W120", "W121", "W122", "W123", "W124", "W125", "W126", "W127", "W128", "W129", "W130", "W131", "W132", "W133", "W134", "W135", "W136", "W137", "W138", "W139", "W140", "W141", "W142", "W143", "W144", "W145", "W146", "W147", "W148", "W149", "W150", "W151", "W152", "W153", "W154", "W155", "W156", "W157", "W158", "W159", "W160", "W161", "W162", "W163", "W164", "W165", "W166", "W167", "W168", "W169", "W170", "W171", "W172", "W173", "W174", "W175", "W176", "W177", "W178", "W179", "W180", "W181", "W182", "W183", "W184", "W185", "W186", "W187", "W188", "W189", "W190", "W191", "W192", "W193", "W194", "W195", "W196", "W197", "W198", "W199", "W200", "W201", "W202", "W203", "W204", "W205", "W206", "W207", "W208", "W209", "W210", "W211", "W212", "W213", "W214", "W215", "W216", "W217", "W218", "W219", "W220", "W221", "W222", "W223", "W224", "W225", "W226", "W227", "W228", "W229", "W230", "W231", "W232", "W233", "W234", "W235", "W236", "W237", "W238", "W239", "W240"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];
% ������Ӧ����
ding = readtable("C:\Users\gyj92\Desktop\2021\C\����1 ��5��402�ҹ�Ӧ�̵��������.xlsx", opts, "UseExcel", false);
%% ת��Ϊ����
ding = table2array(ding);
%% �������
clear  opts
%% ��������
%    ������: C:\Users\gyj92\Desktop\2021\C\����1 ��5��402�ҹ�Ӧ�̵��������.xlsx
%    ������: ��Ӧ�̵Ĺ�������m?��
% �� MATLAB �� 2021-09-10 09:10:59 �Զ�����
%% ����ѡ��
opts = spreadsheetImportOptions("NumVariables", 240);
% ָ����Χ
opts.Sheet = "��Ӧ�̵Ĺ�������m?��";
opts.DataRange = "C2:IH403";
% ָ��������������
opts.VariableNames = ["W001", "W002", "W003", "W004", "W005", "W006", "W007", "W008", "W009", "W010", "W011", "W012", "W013", "W014", "W015", "W016", "W017", "W018", "W019", "W020", "W021", "W022", "W023", "W024", "W025", "W026", "W027", "W028", "W029", "W030", "W031", "W032", "W033", "W034", "W035", "W036", "W037", "W038", "W039", "W040", "W041", "W042", "W043", "W044", "W045", "W046", "W047", "W048", "W049", "W050", "W051", "W052", "W053", "W054", "W055", "W056", "W057", "W058", "W059", "W060", "W061", "W062", "W063", "W064", "W065", "W066", "W067", "W068", "W069", "W070", "W071", "W072", "W073", "W074", "W075", "W076", "W077", "W078", "W079", "W080", "W081", "W082", "W083", "W084", "W085", "W086", "W087", "W088", "W089", "W090", "W091", "W092", "W093", "W094", "W095", "W096", "W097", "W098", "W099", "W100", "W101", "W102", "W103", "W104", "W105", "W106", "W107", "W108", "W109", "W110", "W111", "W112", "W113", "W114", "W115", "W116", "W117", "W118", "W119", "W120", "W121", "W122", "W123", "W124", "W125", "W126", "W127", "W128", "W129", "W130", "W131", "W132", "W133", "W134", "W135", "W136", "W137", "W138", "W139", "W140", "W141", "W142", "W143", "W144", "W145", "W146", "W147", "W148", "W149", "W150", "W151", "W152", "W153", "W154", "W155", "W156", "W157", "W158", "W159", "W160", "W161", "W162", "W163", "W164", "W165", "W166", "W167", "W168", "W169", "W170", "W171", "W172", "W173", "W174", "W175", "W176", "W177", "W178", "W179", "W180", "W181", "W182", "W183", "W184", "W185", "W186", "W187", "W188", "W189", "W190", "W191", "W192", "W193", "W194", "W195", "W196", "W197", "W198", "W199", "W200", "W201", "W202", "W203", "W204", "W205", "W206", "W207", "W208", "W209", "W210", "W211", "W212", "W213", "W214", "W215", "W216", "W217", "W218", "W219", "W220", "W221", "W222", "W223", "W224", "W225", "W226", "W227", "W228", "W229", "W230", "W231", "W232", "W233", "W234", "W235", "W236", "W237", "W238", "W239", "W240"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];
% ��������
gong = readtable("C:\Users\gyj92\Desktop\2021\C\����1 ��5��402�ҹ�Ӧ�̵��������.xlsx", opts, "UseExcel", false);
%% ת��Ϊ����
gong = table2array(gong);
%% �������
clear  opts
%% ��������������
count = zeros(402,1);        %����Υ����������
ding_count = zeros(402,1);      %���ö�����������
xin = zeros(402,1);             %��������
for i = 1:402
    for j = 1:240
        if ding(i,j)~=0         %���������Ϊ�������
            ding_count(i,1) = ding_count(i,1)+1;
            if ding(i,j)~=0&&gong(i,j) ==0        %�������Ϊ�������
                count(i,1) = count(i,1)+1;
            end
        end
    end
    xin(i,1) = count(i,1)/ding_count(i,1);
end

wei = ones(402,1);
for i = 1:402
    if xin(i,1)>0.7
        wei(i,1) = 0;
    end
end

clear c count ding ding_count gong i j ;
