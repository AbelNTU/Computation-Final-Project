function readraw(n)
A=dir(strcat('���d/',num2str(n),'/'));
N=xlsread(strcat('���d/',num2str(n),'/',A(4).name),'2:2');
L=length(N);
   for i=1:L
       x(i)=i-1;
   end
plot(x,N);
end