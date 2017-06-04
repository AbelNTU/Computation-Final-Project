function alldrawtest(o,n,m,l)
% o是公司種類  n是公司代號 m是資料類別 l是資料內的行0數
company=dir('公司種類/');
A=dir(strcat('公司種類/',company(o).name,'/',num2str(n),'/'));
[N,T,data]=xlsread(strcat('公司種類/',company(o).name,'/',num2str(n),'/',A(m).name));
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
