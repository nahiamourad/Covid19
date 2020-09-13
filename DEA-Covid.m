clear;
pkg load io
[Data] = xlsread('COVID19.xlsx','Model 1');%Read data from Excel

ns=7; %number of scenarios
MEff=zeros(ns,size(Data,1),2,2); 
for i=1:ns
CI=4:8;
if(i==1)
CO=1:3;
elseif (i==2)
CO=1;
elseif (i==3)
CO=2;
elseif (i==4)
CO=3;
elseif (i==5)
CO=1:2;
elseif (i==6)
CO=[1,3];
elseif (i==7)
CO=2:3;
endif
        InputD=Data(:,CI)';
        Output=Data(:,CO)';

        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        n=size(Output,2);%number of DMUs
        r=size(Output,1);% number of outputs
        m_D=size(InputD,1);%number of deterministic inputs
        Eff=zeros(n,3);
        
        A=[InputD,zeros(m_D,1)];
        A(m_D+1:m_D+r,:)=[-Output,zeros(r,1)];
        for VRS=0:1
            for O=1:1
                if(O==0)
                    lb=zeros(n+1,1);
                    ub=[Inf(n,1);1];
                elseif(O==1)
                    lb=[zeros(n,1);1];%theta\geq 1
                    ub=Inf(n+1,1);
                end
                for p=1:n
                    A(:,n+1)=zeros(m_D+r,1);
                    if(O==0)
                        A(:,n+1)=[-InputD(:,p);zeros(r,1)];
                        B(:,1)=[zeros(m_D,1);-Output(:,p)];
                    elseif(O==1)
                        A(:,n+1)=[zeros(m_D,1);Output(:,p)];
                        B(:,1)=[InputD(:,p);zeros(r,1)];
                    end
             ctype=repmat('U',[1,size(B,1)]);
             AA=A;
             BB=B;
            if(VRS==1)
                %Sum of lambdas equal one
                AA=[A;ones(1,n),0];
                BB=[B;1];
                ctype=strcat(ctype,'S');%adding inequality constrained
            end
                    f=[zeros(n,1);(-1)^O];%(-1)^O=1 for input oriented and (-1)^O=-1 for output oriented
                    [X,fval] = glpk(f,AA,BB,lb,ub,ctype);% X(n+1)=theta and X(i)=lambda_i
                    %disp(transpose(X));
                    clear AA BB
                    Eff(p,:)=[p,(1-O)*X(n+1)+O*1/X(n+1),X(p)];
                end
                
                MEff(i,:,VRS+1,O+1)=transpose(Eff(:,2));
    end
end
        clearvars -except MEff Data i j ns InputD
end
%%%%%%% Write the results in Excel Sheet

count=1;
n=size(Data,1);
for  i=1:ns
for VRS=0:1 %% CRS and VRS model
    for O=1:1 %% Input and output orientation
        if(VRS==0) && (O==0)
            C(1,count)={"CRS Input Oriented"};
        elseif (VRS==0) && (O==1)
            C(1,count)={"CRS Output Oriented"};
        elseif(VRS==1) && (O==0)
            C(1,count)={"VRS Input Oriented"};
        elseif (VRS==1) && (O==1)
            C(1,count)={"VRS Output Oriented"};            
        endif
        C(2:n+1,count)=num2cell(MEff(i,:,VRS+1,O+1));
        count=count+1;
    end
end
count=count+2;
endfor
xlswrite('Covid19.xlsx',C,'Results','B1');
clearvars -except MEff Data InputD
return