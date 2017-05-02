function allemp()
file=dir('¤ôªd.excel/');
H=length(file)-2;
for i=1:H
    i=i+2;
    a=file(i).name;
    readraw(a,3);
    hold on;
end