function readraw(n,m)
A=dir(strcat('¤ôªd.excel/',num2str(n),'/'));
N=xlsread(strcat('¤ôªd.excel/',num2str(n),'/',A(m).name),'2:2');
L=length(N);
if m==3||m==6
    x = zeros(L);
      for i=1:L
          x(i)=i-1;
      end
   plot(x,N);
else
   L=L/2;
   x = zeros(L);
   for i=1:L
       x(i)=i-1;
       a=2*i-1;
       b(i)=N(a);
   end
   plot(x,b);
   
end   
end