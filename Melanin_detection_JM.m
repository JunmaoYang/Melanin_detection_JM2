% % -------從excel選取量測出的數據-------
wavelength_reflection = xlsread('C:\Users\user\Desktop\楊竣貿\Matlab Program\JM_data.xlsx', 'sheet1');
% wavelength_reflection = xlsread('C:\Users\user2\Desktop\黑色素檢測\test01.xlsx', -1);
Smoothing_1 = zeros(2048,1);
Smoothing_2 = zeros(2048,1);
wavelength = [450;480;500;510;520;530;538;540;542;550;554;560;568;570;576;578;580;590;600;630;660;680;700;750;775;800;805;840;845];
Reflectance_with_oxyhemoglobin = [0.162 0.0672 0.0515 0.0488 0.0598 0.1035 0.139 0.1432 0.1452 0.1201 0.1017 0.0877 0.105 0.1168 0.1526 0.1536 0.1442 0.0426 0.0096 0.0011 0.0008 0.0009 0.0009 0.0014 0.0017 0.002 0.0021 0.0025 0.0025]';
Reflectance_with_deoxyhemoglobin = [0.1351 0.0335 0.0434 0.0538 0.0648 0.0804 0.0994 0.105 0.1109 0.1297 0.1334 0.1309 0.1185 0.1144 0.1007 0.0962 0.0919 0.0687 0.0374 0.0106 0.0081 0.0061 0.0044 0.0039 0.0029 0.002 0.002 0.0019 0.0019]';

Measurement_average = zeros(29,1);
for i = 1:2048
wavelength_reflection(i,2) = wavelength_reflection(i,2)/100;
end
% % % --------複製兩行量測數據--------
for i=1:2048
    Smoothing_1(i,1)= wavelength_reflection(i,2);
    Smoothing_2(i,1)= wavelength_reflection(i,2);
end
%%
% -------去除突起之雜訊-------------
for i = 2 : 2048
     if Smoothing_1(i, 1) > Smoothing_1(i - 1, 1) && Smoothing_1(i, 1) > Smoothing_1(i + 1, 1) && Smoothing_1(i - 1, 1) > Smoothing_1(i + 1, 1)
         Smoothing_2(i, 1) = Smoothing_1(i - 1, 1);
     end
     
     if Smoothing_1(i, 1) > Smoothing_1(i - 1, 1) && Smoothing_1(i, 1) > Smoothing_1(i + 1, 1) && Smoothing_1(i - 1, 1) < Smoothing_1(i + 1, 1)
         Smoothing_2(i, 1) = Smoothing_1(i + 1, 1);
     end
    if Smoothing_1(i, 1) < Smoothing_1(i - 1, 1) && Smoothing_1(i, 1) < Smoothing_1(i + 1, 1) && Smoothing_1(i - 1, 1) < Smoothing_1(i + 1, 1)
        Smoothing_2(i, 1) = Smoothing_1(i - 1, 1);
    end
    
    if Smoothing_1(i, 1) < Smoothing_1(i - 1, 1) && Smoothing_1(i, 1) < Smoothing_1(i + 1, 1) && Smoothing_1(i - 1, 1) > Smoothing_1(i + 1, 1)
        Smoothing_2(i, 1) = Smoothing_1(i + 1, 1);
    end
end

for i = 1:2048
    Smoothing_1(i, 1)=Smoothing_2(i, 1);
end
% % -------計算450-845nm之量測平均------
for j = 1:29
    dn = 0;
    ds = 0;
    ws = wavelength(j,1) - 2;
    wl = wavelength(j,1) + 2;
    for i = 1:2048
        if wavelength_reflection(i,1) > ws && wavelength_reflection(i,1) < wl
            dn = dn + 1;
            ds = ds + Smoothing_1(i,1);
        end
    end
Measurement_average(j,1) = ds / dn;
end
% -------計算660-845nm之量測平均------
for j = 21 : 29
    dn = 0;
    ds = 0;
    ws = wavelength(j,1) - 10;
    wl = wavelength(j,1) + 10;
    for i = 1 : 2048
        if wavelength_reflection(i,1) > ws && wavelength_reflection(i,1) < wl
            dn = dn + 1;
            ds = ds + Smoothing_1(i,1);
        end
    end
Measurement_average(j,1) = ds / dn;
end
x = wavelength;
xi = 450:0.25:845;
Measurement_average_fit = interp1(x,Measurement_average,xi,'pchip');
Reflectance_with_oxyhemoglobin_fit= interp1(x,Reflectance_with_oxyhemoglobin,xi,'pchip');
Reflectance_with_deoxyhemoglobin_fit = interp1(x,Reflectance_with_deoxyhemoglobin,xi,'pchip');
plot(x,Measurement_average,'o',xi,Measurement_average_fit,xi,...
Reflectance_with_oxyhemoglobin_fit,xi,Reflectance_with_deoxyhemoglobin_fit)
%%
%------6係數min及max填入格中---------
worksheet1 = zeros(100000,1);worksheet1(2) = 100;
worksheet2 = zeros(100000,1);worksheet2(2) = 30;
worksheet3 = zeros(100000,1);worksheet3(2) = 100;
worksheet4 = zeros(100000,1);worksheet4(2) = 5;
worksheet5 = zeros(100000,1);worksheet5(2) = 5;
worksheet6 = zeros(100000,1);worksheet6(2) = 5;
worksheet7 = zeros(100000,1);
program_fitting = zeros(29,1);
%%
%--------Montecarlo----------
for f = 1:10
    ns = 100000;          %number of sets: 取6係數的組數
    nd = 99990;         %number of eliminated sets 刪除nd個誤差過大的組
    for kk = 1:8
        %----------取六係數值的範圍-----------------
        comin = min(worksheet1);
        cdmin = min(worksheet2);
        cmmin = min(worksheet3);
        kmin = min(worksheet4);
        umin = min(worksheet5);
        gmin = min(worksheet6);
        
        comax = max(worksheet1);
        cdmax = max(worksheet2);
        cmmax = max(worksheet3);
        kmax = max(worksheet4);
        umax = max(worksheet5);
        gmax = max(worksheet6);
        
        % -------6係數取亂數--------------
        nr = ns ;             %%6係數的最終列"nr"
        ni = nr - nd + 1;     %%給亂數之起始列"ni"
        if kk == 1
            ni = 1;            %%判斷是否第一次使用Monte Carlo
            for i = ni : nr
                worksheet1(i,1) = comin + rand * (comax - comin);
                worksheet2(i,1) = cdmin + rand * (cdmax - cdmin);
                worksheet3(i,1) = cmmin + rand * (cmmax - cmmin);
                worksheet4(i,1) = kmin + rand * (kmax - kmin);
                worksheet5(i,1) = umin + rand * (umax - umin);
                worksheet6(i,1) = gmin + rand * (gmax - gmin);
            end
        end
        %%-------------計算誤差---------------------
        for i = ni:nr
            co = worksheet1(i, 1);  cd = worksheet2(i, 1);  cm = worksheet3(i, 1);
            km = worksheet4(i, 1);  us = worksheet5(i, 1);  gm = worksheet6(i, 1);
            sumdr = 0;   %%差值平方總和
            twl = 0;     %%波長總數
            for j = 1 : 1581
                if xi(1, j) ~= 0
                    wl = xi(1, j);   eo = Reflectance_with_oxyhemoglobin_fit(1, j);
                    ed = Reflectance_with_deoxyhemoglobin_fit(1, j);   mdata = Measurement_average_fit(1, j);
                    dr = mdata - us * (wl / 1000) ^ (-gm) / (0.25 + 0.057 * (co * eo + cd * ed + cm * exp(-km * (wl - 400) / 400)));
                    sumdr = sumdr + dr * dr;
                    twl = twl + 1;
                end
            end
            worksheet7(i, 1) = sqrt(sumdr / twl);
        end
        %-------誤差值由小到大整組排列(注意：範圍)--------
        B = [worksheet1 worksheet2 worksheet3 worksheet4 worksheet5 worksheet6 worksheet7];
        A = sortrows(B,7);
        %-------刪除及縮小6係數的範圍------
        ni = nr - nd + 1;
        for i = ni : nr
            for j = 1:7
                A(i,j) = NaN; worksheet1(i,1) = NaN; worksheet2(i,1) = NaN; worksheet3(i,1) = NaN;...
                    worksheet4(i,1) = NaN; worksheet5(i,1) = NaN; worksheet6(i,1) = NaN; ...
                    worksheet7(i,1) = NaN;
            end
        end
        for i = 1:10
            worksheet1(i,1) = A(i,1); worksheet2(i,1) = A(i,2); worksheet3(i,1) = A(i,3);
            worksheet4(i,1) = A(i,4); worksheet5(i,1) = A(i,5); worksheet6(i,1) = A(i,6);
            worksheet7(i,1) = A(i,7);
        end
        
    end
 end
%%
%--------------Drawing-------------------
co = worksheet1(i, 1);  cd = worksheet2(i, 1);  cm = worksheet3(i, 1);
km = worksheet4(i, 1);  us = worksheet5(i, 1);  gm = worksheet6(i, 1);
for j = 1:1581
    wl = xi(1,j);   eo = Reflectance_with_oxyhemoglobin_fit(1,j);ed = Reflectance_with_deoxyhemoglobin_fit(1,j);
    program_fitting(1,j) = us * (wl / 1000) ^ (-gm) / (0.25 + 0.057 * (co * eo + cd * ed + cm * exp(-km * (wl - 400) / 400)));
end
 %%  
 plot(xi,Measurement_average_fit,'.', xi, program_fitting);
 xlabel('wavelength,nm');ylabel('reflection,%')
 legend('original','program_ fitting');
 
 for i = 1:length(xi)
data{i,1}=xi(i);
data{i,2}=Measurement_average_fit(i);

end
for i = 1:length(xi)
data{i,3}=xi(i);
data{i,4}=program_fitting(i);

end
[status, message] = xlswrite('output.xlsx', data,'sheet2');











