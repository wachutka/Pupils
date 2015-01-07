% Import pupil data from excel files to matlab 
function [NormI1, NormP1, NormI2, NormP2, NormI3, NormP3] = Import_Pupils(subject,version)

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Must specify input file name here!  In the format of VX_BY_##.xls where 
% X = version number (A=1, B=2, C=3), Y = block number (1:3), and ## is the
% subject number.

%InputFile = 'V1_B3_27.xls';    % Input data file from Gazetracker (must be exported from Gazetracker as .xls)
for block = 1:3
    version
    subject
    InputFile = ['V' num2str(version) '_B' num2str(block) '_' num2str(subject) '.xls']
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



% Times of target word onset for each of the audio files

Times = [3.76900000000000,3.46900000000000,3.63500000000000,4.20300000000000,3.33500000000000,3.83600000000000,3.93600000000000,3.80200000000000,4.16900000000000,3.83600000000000,3.90200000000000,3.90200000000000,3.90200000000000,4.06900000000000,4.00300000000000,3.86900000000000,4.67000000000000,4.40300000000000,4.60300000000000,3.30200000000000,3.20200000000000,4.10300000000000,2.73500000000000,3.60200000000000,3.20200000000000,3.40200000000000,3.16800000000000,3.70200000000000,2.93500000000000,3.40200000000000,3.80200000000000,4.03600000000000,4.26900000000000,3.86900000000000,3.83600000000000,4.30300000000000,3.33500000000000,3.20200000000000,3.26800000000000,3.16800000000000,3.63500000000000,3.83600000000000,2.90100000000000,3.40200000000000,3.70200000000000,3.70200000000000,3.66900000000000,3.86900000000000,3.10200000000000,4.77000000000000,3.80200000000000,3.40200000000000,4.40300000000000,3.96900000000000,4.50300000000000,4.33600000000000,3.96900000000000,3.33500000000000,3.33500000000000,3.96900000000000,4.43600000000000,3.70200000000000,3.30200000000000,3.30200000000000,4.50300000000000,3.96900000000000];
   
   
% Individual worksheets must be imported one at a time; worksheet names must be
% strings due to MATLAB's 'basic' limitations when importing from excel on
% a mac.  Hence these ridiculously long lines of worksheet names. 

% For Version A
Worksheets1 = { '3-White+.wmv - G' '4-Black+.wmv - G' '13-24_1.wmv - G' '15-22_2.wmv - G' '17-51_1.wmv - G' '19-62_1.wmv - G' '21-16_1.wmv - G' '23-19_2.wmv - G' '25-23_1.wmv - G' '27-12_2.wmv - G' '30-25_1.wmv - G' '32-26_2.wmv - G' '34-44_1.wmv - G' '36-21_1.wmv - G' '38-50_2.wmv - G' '40-37_1.wmv - G' '42-58_1.wmv - G' '45-10_2.wmv - G' '47-34_1.wmv - G' '49-43_1.wmv - G' '51-32_2.wmv - G' '53-46_1.wmv - G' '55-1_1.wmv - G' };

Worksheets2 = { '2-17_1.wmv - G' '4-40_1.wmv - G' '6-57_1.wmv - G' '8-49_1.wmv - G' '10-14_2.wmv - G' '13-9_2.wmv - G' '15-33_1.wmv - G' '17-20_1.wmv - G' '19-41_1.wmv - G' '21-48_1.wmv - G' '23-2_1.wmv - G' '25-7_2.wmv - G' '27-60_2.wmv - G' '30-35_2.wmv - G' '32-18_2.wmv - G' '34-36_1.wmv - G' '36-56_1.wmv - G' '38-61_1.wmv - G' '40-47_1.wmv - G' '42-54_2.wmv - G' '44-13_1.wmv - G' };

Worksheets3 = { '2-11_1.wmv - G' '4-30_1.wmv - G' '6-8_1.wmv - G' '8-63_2.wmv - G' '10-27_2.wmv - G' '12-55_2.wmv - G' '15-29_2.wmv - G' '17-3_2.wmv - G' '19-38_2.wmv - G' '21-53_1.wmv - G' '23-52_2.wmv - G' '25-39_2.wmv - G' '27-5_2.wmv - G' '30-4_2.wmv - G' '32-65_1.wmv - G' '34-15_2.wmv - G' '36-6_2.wmv - G' '38-31_1.wmv - G' '40-28_2.wmv - G' '42-59_2.wmv - G' '44-45_1.wmv - G' };

% For Version B
Worksheets4 = { '3-White+.wmv - G' '4-Black+.wmv - G' '13-18_1.wmv - G' '15-7_1.wmv - G' '17-43_2.wmv - G' '19-1_2.wmv - G' '21-24_2.wmv - G' '23-51_2.wmv - G' '25-46_2.wmv - G' '27-54_1.wmv - G' '30-34_2.wmv - G' '32-23_2.wmv - G' '34-16_2.wmv - G' '36-37_2.wmv - G' '38-58_2.wmv - G' '40-21_2.wmv - G' '42-9_1.wmv - G' '45-44_2.wmv - G' '47-14_1.wmv - G' '49-62_2.wmv - G' '51-35_1.wmv - G' '53-25_2.wmv - G' '55-60_1.wmv - G' };

Worksheets5 = { '2-22_2.wmv - G' '4-6_1.wmv - G' '6-55_1.wmv - G' '8-39_1.wmv - G' '10-26_2.wmv - G' '13-63_1.wmv - G' '15-19_2.wmv - G' '17-28_1.wmv - G' '19-59_1.wmv - G' '21-29_1.wmv - G' '23-15_1.wmv - G' '25-3_1.wmv - G' '27-5_1.wmv - G' '30-10_2.wmv - G' '32-12_2.wmv - G' '34-32_2.wmv - G' '36-4_1.wmv - G' '38-52_1.wmv - G' '40-38_1.wmv - G' '42-50_2.wmv - G' '44-27_1.wmv - G' };

Worksheets6 = { '2-2_1.wmv - G' '4-49_1.wmv - G' '6-11_2.wmv - G' '8-20_1.wmv - G' '10-56_1.wmv - G' '12-47_1.wmv - G' '15-8_2.wmv - G' '17-57_1.wmv - G' '19-33_1.wmv - G' '21-31_2.wmv - G' '23-48_1.wmv - G' '25-61_1.wmv - G' '27-65_2.wmv - G' '30-17_1.wmv - G' '32-45_2.wmv - G' '34-41_1.wmv - G' '36-13_1.wmv - G' '38-36_1.wmv - G' '40-30_2.wmv - G' '42-53_2.wmv - G' '44-40_1.wmv - G' };

% For Version C
Worksheets7 = { '3-White+.wmv - G' '4-Black+.wmv - G' '13-31_2.wmv - G' '15-1_1.wmv - G' '17-21_1.wmv - G' '19-25_1.wmv - G' '21-23_1.wmv - G' '23-16_1.wmv - G' '25-30_2.wmv - G' '27-11_2.wmv - G' '30-34_1.wmv - G' '32-43_1.wmv - G' '34-46_1.wmv - G' '36-45_2.wmv - G' '38-37_1.wmv - G' '40-58_1.wmv - G' '42-51_1.wmv - G' '45-8_2.wmv - G' '47-24_1.wmv - G' '49-53_2.wmv - G' '51-62_1.wmv - G' '53-44_1.wmv - G' '55-65_2.wmv - G' };

Worksheets8 = { '2-10_1.wmv - G' '4-48_2.wmv - G' '6-22_1.wmv - G' '8-56_2.wmv - G' '10-13_2.wmv - G' '13-17_2.wmv - G' '15-32_1.wmv - G' '17-12_1.wmv - G' '19-57_2.wmv - G' '21-49_2.wmv - G' '23-40_2.wmv - G' '25-47_2.wmv - G' '27-33_2.wmv - G' '30-19_1.wmv - G' '32-2_2.wmv - G' '34-26_1.wmv - G' '36-61_2.wmv - G' '38-50_1.wmv - G' '40-20_2.wmv - G' '42-41_2.wmv - G' '44-36_2.wmv - G' };

Worksheets9 = { '2-63_1.wmv - G' '4-35_2.wmv - G' '6-59_1.wmv - G' '8-3_1.wmv - G' '10-29_1.wmv - G' '12-55_1.wmv - G' '15-15_1.wmv - G' '17-18_2.wmv - G' '19-60_2.wmv - G' '21-28_1.wmv - G' '23-27_1.wmv - G' '25-5_1.wmv - G' '27-14_2.wmv - G' '30-54_2.wmv - G' '32-9_2.wmv - G' '34-38_1.wmv - G' '36-6_1.wmv - G' '38-52_1.wmv - G' '40-7_2.wmv - G' '42-39_1.wmv - G' '44-4_1.wmv - G' };

j = 1;  % Counter for bins
k = 1;  % Counter for bins

%%%%%
if findstr(InputFile, 'B1') 
    
        nwkshts = length(Worksheets1);
        B = 1;  % Block identifier for white/black slides
        
else
        nwkshts = length(Worksheets2);
        B = 0;
end

DataAll = zeros(1000,nwkshts*2);    % Each slide has 2 columns - 1 for pupil dia 1 for time

for i = 1:nwkshts        
    
    if findstr(InputFile, 'V1_B1') 
        
        wks = char(Worksheets1(i));     % Identify current worksheet
        Raw = xlsread(InputFile,wks);   % Read .xls file
        Data = 'V1_B1';
     
    elseif findstr(InputFile, 'V1_B2')
        
        wks = char(Worksheets2(i));       % Identify current worksheet
        Raw = xlsread(InputFile,wks);     % Read .xls file
        Data = 'V1_B2';
       
    elseif findstr(InputFile, 'V1_B3')
        
        wks = char(Worksheets3(i));       % Identify current worksheet
        Raw = xlsread(InputFile,wks);     % Read .xls file
        Data = 'V1_B3';
        
    elseif findstr(InputFile, 'V2_B1') 
        
        wks = char(Worksheets4(i));     % Identify current worksheet
        Raw = xlsread(InputFile,wks);   % Read .xls file
        Data = 'V2_B1';
     
    elseif findstr(InputFile, 'V2_B2')
        
        wks = char(Worksheets5(i));       % Identify current worksheet
        Raw = xlsread(InputFile,wks);     % Read .xls file
        Data = 'V2_B2';
       
    elseif findstr(InputFile, 'V2_B3')
        
        wks = char(Worksheets6(i));       % Identify current worksheet
        Raw = xlsread(InputFile,wks);     % Read .xls file
        Data = 'V2_B3';
        
    elseif findstr(InputFile, 'V3_B1') 
        
        wks = char(Worksheets7(i));     % Identify current worksheet
        Raw = xlsread(InputFile,wks);   % Read .xls file
        Data = 'V3_B1';
     
    elseif findstr(InputFile, 'V3_B2')
        
        wks = char(Worksheets8(i));       % Identify current worksheet
        Raw = xlsread(InputFile,wks);     % Read .xls file
        Data = 'V3_B2';
       
    elseif findstr(InputFile, 'V3_B3')
        
        wks = char(Worksheets9(i));       % Identify current worksheet
        Raw = xlsread(InputFile,wks);     % Read .xls file
        Data = 'V3_B3';
        
    end
     
     sz = size(Raw,1);              % Size of data for current slide
     
     DataAll(1:sz,(i*2-1)) = Raw(:,3);      % Create matrix with all data combined
                                            % *2 for 2 columns 
     DataAll(1:sz,(i*2)) = Raw(:,4);
     
     Raw = Raw(:,3:4);              % Eliminate unnecessary columns
    
     %%%%% This corrects for camera thinking reflection is pupil %%%%%
     
     R = Raw(:,1);      % Select all pupil size data
     R(R<40) = NaN;     % Remove data from reflections (not actual pupil size)
     Raw(:,1) = R;      % Return corrected data to Raw
     
     %%%%%
     
    s = sscanf(wks(strfind(wks,'-')+1:strfind(wks,'-')+2),'%d');
        
    t = Times(s);
   
    if B == 1 && i == 1 | i == 2
        t = 10;
    end
    
    t = t+0.0001;      % Time of target word onset +0.0001 to choose later sample if equil distance between two samples
         
    Raw(:,3) = abs(Raw(:,2)-t);   % Calculate difference between target onset and sample time
 
    Target(i) = find(Raw(:,3)==min(Raw(:,3)));  % Find target word onset point
    
   if B == 1 && i == 1 | i == 2
        
       Range(i,1) = nanmean(Raw(Target(i):end,1));  % Means for black and white screens
        
   else  
           
       if sscanf(wks(strfind(wks,'_')+1),'%d') == 1    % Probable

           if length(Raw(Target(i):end,1)) > 180
           
                Prob(j,1) = nanmean(Raw(Target(i)-12:Target(i)-1,1));    % Baseline
                Prob(j,2) = nanmean(Raw(Target(i):Target(i)+12,1));      % First 12 samples
                Prob(j,3) = nanmean(Raw(Target(i)+13:Target(i)+24,1));   % Second 12 samples
                Prob(j,4) = nanmean(Raw(Target(i)+25:Target(i)+36,1));   % Third 12 samples etc...
                Prob(j,5) = nanmean(Raw(Target(i)+37:Target(i)+48,1));
                Prob(j,6) = nanmean(Raw(Target(i)+49:Target(i)+60,1));
                Prob(j,7) = nanmean(Raw(Target(i)+61:Target(i)+72,1));
                Prob(j,8) = nanmean(Raw(Target(i)+73:Target(i)+84,1));
                Prob(j,9) = nanmean(Raw(Target(i)+85:Target(i)+96,1));
                Prob(j,10) = nanmean(Raw(Target(i)+97:Target(i)+108,1));
                Prob(j,11) = nanmean(Raw(Target(i)+109:Target(i)+120,1));
                Prob(j,12) = nanmean(Raw(Target(i)+121:Target(i)+132,1));
                Prob(j,13) = nanmean(Raw(Target(i)+133:Target(i)+144,1));
                Prob(j,14) = nanmean(Raw(Target(i)+145:Target(i)+156,1));
                Prob(j,15) = nanmean(Raw(Target(i)+157:Target(i)+168,1));
                Prob(j,16) = nanmean(Raw(Target(i)+169:Target(i)+180,1));
                
            
                PeakP(j) = max(Raw(Target(i)+1:Target(i)+180,1)); % Peak pupil size
                LocPkP = find(Raw(Target(i)+1:Target(i)+180,1) == PeakP(j));   % Locate peak
                if numel(LocPkP) >= 1  
                    FirstPeakP = LocPkP(1);     % First peak
                    LPeakP(j) = Raw(Target(i)+FirstPeakP,2)-t; % Latency to peak pupil size
                else
                    LPeakP(j) = NaN;
                end
                
           else                  % If not enough data exists in file
                PeakP(j) = NaN;
                LPeakP(j) = NaN;
                Prob(j,1:16) = NaN;
               
           end
           
            j = j + 1;
            
       elseif sscanf(wks(strfind(wks,'_')+1),'%d') == 2         % Improbable

           if length(Raw(Target(i):end,1)) > 180
           
                Impr(k,1) = nanmean(Raw(Target(i)-12:Target(i)-1,1));    % Baseline
                Impr(k,2) = nanmean(Raw(Target(i):Target(i)+12,1));      % First 12 samples
                Impr(k,3) = nanmean(Raw(Target(i)+13:Target(i)+24,1));   % Second 12 samples
                Impr(k,4) = nanmean(Raw(Target(i)+25:Target(i)+36,1));   % Third 12 samples etc...
                Impr(k,5) = nanmean(Raw(Target(i)+37:Target(i)+48,1));
                Impr(k,6) = nanmean(Raw(Target(i)+49:Target(i)+60,1));
                Impr(k,7) = nanmean(Raw(Target(i)+61:Target(i)+72,1));
                Impr(k,8) = nanmean(Raw(Target(i)+73:Target(i)+84,1));
                Impr(k,9) = nanmean(Raw(Target(i)+85:Target(i)+96,1));
                Impr(k,10) = nanmean(Raw(Target(i)+97:Target(i)+108,1));
                Impr(k,11) = nanmean(Raw(Target(i)+109:Target(i)+120,1));
                Impr(k,12) = nanmean(Raw(Target(i)+121:Target(i)+132,1));
                Impr(k,13) = nanmean(Raw(Target(i)+133:Target(i)+144,1));
                Impr(k,14) = nanmean(Raw(Target(i)+145:Target(i)+156,1));
                Impr(k,15) = nanmean(Raw(Target(i)+157:Target(i)+168,1));
                Impr(k,16) = nanmean(Raw(Target(i)+169:Target(i)+180,1));

                PeakI(k) = max(Raw(Target(i)+1:Target(i)+180,1)); % Peak pupil size
                LocPkI = find(Raw(Target(i)+1:Target(i)+180,1) == PeakI(k));   % Locate peak
                if numel(LocPkI) >= 1
                    FirstPeakI = LocPkI(1);     % First peak
                    LPeakI(k) = Raw(Target(i)+FirstPeakI,2)-t; % Latency to peak pupil size
                else
                    LPeakI(k) = NaN;
                end
           else                  % If not enough data exists in file
               
               PeakI(k) = NaN;
               LPeakI(k) = NaN;
               Impr(k,1:16) = NaN;
               
           end
           
            k = k + 1;
            
       end
        
   end

end

LPeakP(LPeakP<0)=0;     % Correct for possible negative values when peak is the first sample
LPeakI(LPeakI<0)=0;

%%%%%%% Normalize to baseline and find max pupil size %%%%%%%

for c = 1:size(Prob,1)
    for v = 2:length(Prob)
        NormP(c,v-1) = Prob(c,v)/Prob(c,1);   
    end
    
    NormP(c,length(Prob)) = PeakP(c)/Prob(c,1);  % Peak pupil size
    NormP(c,length(Prob)+1) = LPeakP(c)*1000;         % Latency to peak pupil size from target word onset
end

 
for C = 1:size(Impr,1)
    for V = 2:length(Impr)
        NormI(C,V-1) = Impr(C,V)/Impr(C,1); 
    end 
    
    NormI(C,length(Impr)) = PeakI(C)/Impr(C,1);  % Peak pupil size
    NormI(C,length(Impr)+1) = LPeakI(C)*1000;         % Latency to peak pupil size from target word onset
end
   
if block == 1
    NormI1(1:14,1:17) = NaN;
    NormP1(1:14,1:17) = NaN;
    NormI1(1:size(NormI,1),1:size(NormI,2))= NormI;
    NormP1(1:size(NormP,1),1:size(NormP,2)) = NormP;
else if block == 2
    NormI2(1:14,1:17) = NaN;
    NormP2(1:14,1:17) = NaN;
    NormI2(1:size(NormI,1),1:size(NormI,2)) = NormI;
    NormP2(1:size(NormP,1),1:size(NormP,2)) = NormP;   
    else if block == 3
            NormI3(1:14,1:17) = NaN;
            NormP3(1:14,1:17) = NaN;
            NormI3(1:size(NormI,1),1:size(NormI,2)) = NormI;
            NormP3(1:size(NormP,1),1:size(NormP,2)) = NormP;
        end
    end
end
% NormI1
% NormP1
% NormI2
% NormP2
% NormI3
% NormP3
end
xlswrite('Pupil Data Older Adults.xlsx',NormI1,2,'B3:R16')
xlswrite('Pupil Data Older Adults.xlsx',NormP1,2,'T3:AJ16')
xlswrite('Pupil Data Older Adults.xlsx',NormI2,2,'B21:R34')
xlswrite('Pupil Data Older Adults.xlsx',NormP1,2,'T21:AJ34')
xlswrite('Pupil Data Older Adults.xlsx',NormI3,2,'B39:R52')
xlswrite('Pupil Data Older Adults.xlsx',NormP3,2,'T39:AJ52')
%NormMeanMaxProb = mean(NormP(:,length(Prob)))
%NormMeanMaxImpr = mean(NormI(:,length(Impr)))

