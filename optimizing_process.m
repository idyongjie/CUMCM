gong = xlsread("gong.xlsx");    %获取供应数据
%% 获得290家企业序号
xuhao = [140
151
229
348
374
361
108
126
330
201
139
308
307
282
275
329
356
268
306
395
340
143
37
194
284
352
131
247
365
294
80
218
244
86
114
266
150
123
7
314
31
291
338
74
3
364
40
55
367
346
76
189
129
178
53
210
78
239
273
152
237
245
113
66
75
25
5
23
186
269
64
263
110
115
208
292
154
157
318
98
221
304
253
102
149
106
270
133
271
92
33
65
362
267
357
265
331
146
39
332
381
165
383
342
379
128
300
202
11
392
227
197
30
360
163
175
324
16
35
216
138
258
159
174
256
21
279
226
301
172
349
122
359
141
206
274
88
313
296
99
191
209
87
90
27
326
336
54
26
70
13
207
170
116
377
41
254
376
96
89
91
184
182
325
32
42
217
205
20
61
107
310
52
132
312
380
173
341
171
167
242
220
94
121
255
145
82
391
252
250
347
384
148
223
72
354
195
370
363
97
83
104
85
368
17
101
278
243
56
169
398
127
369
185
345
136
158
155
320
288
60
196
187
323
333
109
281
401
36
69
280
225
233
388
117
386
235
193
24
389
50
81
230
286
353
188
316
232
199
399
272
327
298
73
259
48
105
249
319
10
322
18
120
261
358
79
401
311
124
303
200
4
118
29
47
378
315
248
111
371
57
28
335
147
264
134
222
142
166
112
];
%% 将290家企业分类  求数据
[hang,~] = size(xuhao);
GA = [];
GB = [];
GC = [];
for i = 1:hang
    if gong(xuhao(i),2) == 1
        GA=[GA;gong(xuhao(i),:)];
    elseif gong(xuhao(i),2) == 2
        GB=[GB;gong(xuhao(i),:)];
    else    gong(xuhao(i),2) == 3
            GC=[GC;gong(xuhao(i),:)];
    end
end
clear hang;     
%% 获取五年数据
[ga,~] = size(GA);
[gb,~] = size(GB);
[gc,~] = size(GC);
GAm = zeros(ga,49);
GAm(:,1) = GA(:,1);
GBm = zeros(gb,49);
GBm(:,1) = GB(:,1);
GCm = zeros(gc,49);
GCm(:,1) = GC(:,1);
for i = 1:ga
    for j = 1:48
        GAm(i,j+1) = max([GA(i,2+j),GA(i,50+j),GA(i,98+j),GA(i,146+j),GA(i,194+j)]);
    end
end
for i = 1:gb
    for j = 1:48
        GBm(i,j+1) = max([GB(i,2+j),GB(i,50+j),GB(i,98+j),GB(i,146+j),GB(i,194+j)]);
    end
end
for i = 1:gc
    for j = 1:48
        GCm(i,j+1) = max([GC(i,2+j),GC(i,50+j),GC(i,98+j),GC(i,146+j),GC(i,194+j)]);
    end
end
clear GA GB GC i j i ga gb gc;
%% 获得储存矩阵
chu = zeros(1,48);
he = sum (GAm/0.6)+sum(GBm/0.66)+sum(GCm/0.72);
he(1) = [];
xuqiu = 28200*ones(1,48);
xuqiu(1) = xuqiu(1)+56400;
he = he-xuqiu;
for i = 1:48
    chu(i) = sum(he(1:i));
end
%% 给序号加分类
qiye = zeros(290,1);
for i =1:290
    qiye(i) = gong(xuhao(i),2);
end
xuhao = [xuhao,qiye];
clear qiye;
%% 循环剔除
num = 290;
game = 1;
[a,~]=size(GAm);
[b,~]=size(GBm);
[c,~]=size(GCm);
lin = [];
while game == 1
    if xuhao(num,2) == 1
        lin = GAm(a,2:49)/0.6;
        a=a-1;
        num=num-1;
    end
    if xuhao(num,2) == 2
        lin = GBm(b,2:49)/0.66;
        b=b-1;
        num=num-1;
    end
    if xuhao(num,2) == 3
        lin = GCm(c,2:49)/0.72;
        c=c-1;
        num=num-1;
    end
    for i = 1:48
        he(i) = sum(lin(1:i));

    end
    chu = chu-he;
    for i = 1:48
        if chu(i)<0 || num == 0
            game = 0;
        end
    end
end
%% 寻求是否还有其他能替换的企业
num = num+1;
if xuhao(num,2) == 1
    a=a+1;
    chu(1,:) = chu(1,:)+GAm(a,2:49);
elseif xuhao(num,2) == 2
    b=b+1;
    chu(1,:) = chu(1,:)+GBm(b,2:49);
elseif xuhao(num,2) == 3
    c=c+1;
    chu(1,:) = chu(1,:)+GCm(c,2:49);
end
qianshuju = [GAm(1:a,2:49)/0.6;GBm(1:b,2:49)/0.66;GCm(1:c,2:49)/0.72];
lin = [GAm(1:a,1);GBm(1:b,1);GCm(1:c,1);];
qianshuju = [lin,qianshuju];
clear lin;
qsj = zeros(a+b+c,49);
qsj(:,1) = qianshuju(:,1);
for i = 2:49
    for j =1:num
     qsj(j,i) = sum(qianshuju(j,2:i));
    end
end
%% excel 处理数据，再导入表格
pai = xlsread("pai1.xlsx");
[m,~] = size(pai);
jieguo = pai(1:m-1,:);
chucun = pai(m,2:49);
clear a b c chu GAm game GBm GCm gong he i j lin m num pai qianshuju qsj;
%% 优化
[zhi,n]=min(chucun);
youhua = zeros(1,48);
chu=1;
jilu = 0;
for i=1:48
    for i = chu:48
        chucun(i) = chucun(i)-zhi;
        youhua(i) = youhua(i)+zhi;
    end
    [zhi,n]=min(chucun(jilu+1:48));
    chu = jilu+1;
    jilu = jilu+n;
    if jilu>24
        break;
    end
end
%% 处理优化数据
A = xlsread("pai1.xlsx");
youhua1 = zeros(17,49);
youhua1 = A(1:17,1:49);
juzhen = youhua1;
for i =49:-1:3
    juzhen(:,i)=youhua1(:,i)-youhua1(:,i-1);
end
youhua1 = juzhen;
clear juzhen;
for i = 48:-1:2
    youhua(i) = youhua(i)-youhua(i-1);
end
for j = 1:48
    lin = sum(youhua1(1:17,1+j));
    for i = 1:17
        youhua1(i,j+1) = youhua1(i,j+1)-youhua1(i,j+1)*youhua(j)/lin;
    end
end
%% 信誉
xinyu = [0.004545455
0
0
0.255813953
0
0
0
0
0.005813953
0
0
0
0
0
0
0.025125628
0.428571429
];
%% 求订货方案
sun = xlsread("yuan.xlsx");
youhua2 = youhua1;
for j = 2:49
    for i = 1:17
        youhua2(i,j) = youhua1(i,j)/((1-xinyu(i))*0.9*(1-0.01*sun(j-1)));
    end
end
%% 转化为材料
gong=xlsread("gong.xlsx");
for i=1:17
    if gong(youhua2(i),2)==1
        youhua2(i,2:49) = youhua2(i,2:49)*0.6;
    elseif gong(youhua2(i),2)==2
        youhua2(i,2:49) = youhua2(i,2:49)*0.66;
    else gong(youhua2(i),2)==3
        youhua2(i,2:49) = youhua2(i,2:49)*0.72;
    end
end
X = 9999*ones(402,49);
X(:,1) = 1:402;
for i =1:17
    X(youhua2(i,1),:) = youhua2(i,:);
end