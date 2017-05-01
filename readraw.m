function readraw(n,m)
A=dir(strcat('¤ôªd/',num2str(n),'/'));
N=xlsread(strcat('¤ôªd/',num2str(n),'/',A(m).name),'2:2');
L=length(N);
if m==4|m==9
      for i=1:L
          x(i)=i-1;
      end
   plot(x,N);
else
   g=L/2;
   for i=1:g
       x(i)=i-1;
       a=2*i-1
       b(i)=N(a);
   end
   plot(x,b);
   
end   
end