function alldraw(n,m,l)
A=dir(strcat('¤ôªd.excel/',num2str(n),'/'));
[N,T,data]=xlsread(strcat('¤ôªd.excel/',num2str(n),'/',A(m).name));
[e,f]=size(T);
f=f-1;
if m==3||m==6
    x = zeros(f);
      for i=1:f
          x(i)=i-1;
          y(i)=N(l-1,i);
      end
   plot(x,y);
else
   f=f/2;
   x = zeros(f);
   for i=1:f
       x(i)=i-1;
       a=2*i-1;
       y(i)=N(l-1,a);
   end
   plot(x,y);   
end   
end