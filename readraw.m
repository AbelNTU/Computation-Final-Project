function readraw()
N=xlsread('1101(102.1~103.4)EPM.xls','2:2');
L=length(N);
   for i=1:L
       x(i)=i-1;
   end
plot(x,N);
end