%% �����ȷ���
clear
x= xlsread("fuben.xlsx");
% ��A���Ż�
b=x;
figure;
he = [];
    for k = 1:24
        he(k) = sum(x(:,k+2));
    end
for i = 0.1:0.01:0.2
    for j = 1:18
        %�ı�����x��ֵ�����л�����
        if x(j,2) == 1
            b(j,2:26) = x(j,2:26)/(1-i);
        end
    end
    lin =[];
    for k = 1:24
        lin(k) = sum(b(:,k+2));
    end
    hold on;
    zeng = 100*(lin-he)./he;
    plot((zeng));
end
legend("10%","11%","12%","13%","14%","15%","16%","17%","18%","19%","20%")
xlabel("�ܴ�");
ylabel("������%��");