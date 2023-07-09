import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import os
import pyecharts.options as opts
from pyecharts.charts import Bar, Line, Grid, Pie,Tab,Radar,Kline,WordCloud
import xlwings as xw
matplotlib.use('Agg')
import time
import jieba
import jieba.analyse
from wordcloud import WordCloud as WD
import difflib

pd.options.mode.chained_assignment = None

plt.rcParams['font.sans-serif'] = ['Songti SC']
plt.rcParams['axes.unicode_minus'] = False

trade_date = '交易日期'
trade_time = '交易时间'
trade_in = '贷方金额(收入)' 
trade_out = '借方金额(支出)'
left_amount = '余额'
trade_player = '对方户名|单位|收(付)方名称'
trade_abstract = ['备注','摘要','用途','用途[ Purpose ]']

#TYPE B
outcome_amounnt = ['发生额', '交易金额[ Trade Amount ]']
trade_direction = ['交易方向','资金流向','交易类型[ Transaction Type ]']
Payee = ["收款人名称[ Payee's Name ]"]
Payer = ["付款人名称[ Payer's Name ]"]


## 合并各个银行的流水为统一的Dataframe
def creamfind(vector,columns):
    for i in range(len(vector)):
        res = difflib.get_close_matches(vector[i], columns, 1, cutoff=0.6)
        if res != []:
            break
    return res

    
def mergeflow(filefolder='../upload'):

    rootpath = filefolder
    files = os.listdir(rootpath)
    all_df = []
    head_col = 0
    begin_amount = 0
    for path in files:
        print("正在读取:"+path)
        head_col = 0
        if os.path.splitext(path)[-1] in ['.xls', '.xlsx', '.csv']:
            df = pd.read_excel(rootpath+'/'+path)
            if 'Unnamed: 2' in df.columns.values:
                for i in range(len(df)):
                    if df.iloc[i].isnull().any() == False:
                        head_col = i+1
                        break
            df = pd.read_excel(rootpath+'/'+path, header=head_col)
        else:
            continue

        print('原始列名:', df.columns.values)
        column_names = np.array(df.columns, dtype='str')
        direction = creamfind(trade_direction, column_names)
        outcome = creamfind(outcome_amounnt, column_names)
        abstract = creamfind(trade_abstract,column_names)


        date = difflib.get_close_matches(
            trade_date, column_names, 1, cutoff=0.1)
        time = difflib.get_close_matches(
            trade_time, column_names, 1, cutoff=0.1)
        amount = difflib.get_close_matches(
            left_amount, column_names, 1, cutoff=0.1)
        player = difflib.get_close_matches(
            trade_player, column_names, 1, cutoff=0.1)
        # abstract = difflib.get_close_matches(
        #     trade_abstract, column_names, 1, cutoff=0.1)
        
        
        if direction != [] and outcome != []:  # TypeB
            print('该流水非区分贷方与借方，将自动进行区分')
            temp_in = []
            temp_out = []
            for i in range(len(df)):
                if df[direction].iloc[i].values in ['入账', '贷方', '来账']:
                    temp_in.append(np.abs(df[outcome].iloc[i].values[0]))
                    temp_out.append(0)
                else:
                    temp_in.append(0)
                    temp_out.append(np.abs(df[outcome].iloc[i].values[0]))
            df[trade_in] = temp_in
            df[trade_out] = temp_out
            
            payee = creamfind(Payee, column_names)
            payer = creamfind(Payer, column_names)
            print(payee,payer)
            if payee != [] and payer != []:
                temp_player = []
                for i in range(len(df)):
                    if df[trade_in].iloc[i]!=0:
                        print(df[payee].iloc[i])
                        temp_player.append(df[payee].iloc[i].values[0])
                    else:
                        temp_player.append(df[payer].iloc[i].values[0])
                df[trade_player] = temp_player
                col = [date[0], time[0], trade_in, trade_out,amount[0], trade_player, abstract[0]]
            else:
                col = [date[0], time[0], trade_in, trade_out,amount[0], player[0], abstract[0]]

        else:  # Type A
            t_in = difflib.get_close_matches(
                trade_in, column_names, 1, cutoff=0.1)
            t_out = difflib.get_close_matches(
                trade_out, column_names, 1, cutoff=0.1)
            df[trade_in] = df[t_in[0]]
            df[trade_out] = df[t_out[0]]
            col = [date[0], time[0], t_in[0], t_out[0],amount[0], player[0], abstract[0]]
        print("经过匹配挑选出来的列表名：",col)
        
        df = df[col]
        real_col = [trade_date, trade_time, trade_in, trade_out,
                    left_amount, trade_player, trade_abstract[0]]
        print("上述列表重命名为：",real_col)
        df.columns = real_col

        df[trade_time] = df[trade_time].astype(str)
        df[trade_time] = df[trade_time].str.zfill(6)
        df[trade_date] = df[trade_date].astype(str)
        df[left_amount] = df[left_amount].replace(',', '', regex=True)
        df[left_amount] = df[left_amount].astype(float)
        df[trade_out] = df[trade_out].replace(',', '', regex=True)
        df[trade_out] = df[trade_out].astype(float)
        df[trade_in] = df[trade_in].replace(',', '', regex=True)
        df[trade_in] = df[trade_in].astype(float)
        df[trade_abstract[0]] = df[trade_abstract[0]].fillna('')
        df[trade_abstract[0]] = df[trade_abstract[0]].astype(str)
        
        

        df = df.dropna(subset=[left_amount])
        df[trade_in] = df[trade_in].fillna(0)
        df[trade_out] = df[trade_out].fillna(0)

        if len(df[trade_time].iloc[0]) <= 10 and df[trade_time].iloc[0] != df[trade_date].iloc[0]:
            df[trade_time] = df[trade_date] + ' ' + df[trade_time]
        df[trade_time] = pd.to_datetime(df[trade_time])
        df = df.sort_values(by=[trade_time], kind='merge')
        begin_amount += (df[left_amount].iloc[0] -
                         df[trade_in].iloc[0]+df[trade_out].iloc[0])

        all_df.append(df)
        print('\n')

    df = pd.concat(all_df)

    df = df.sort_values(by=[trade_time], kind='merge')
    df.drop_duplicates(subset=None, keep='first', inplace=True)

    df[left_amount].iloc[0] = begin_amount

    for i in range(len(df)):
        temp = df[left_amount].iloc[i] + \
            df[trade_in].iloc[i]-df[trade_out].iloc[i]
        df[left_amount].iloc[i] = temp
        if i == len(df)-1:
            break
        else:
            df[left_amount].iloc[i+1] = temp

    df = df.drop([trade_date], axis=1)
    df.reset_index(drop=True, inplace=True)

    for i in range(len(df)):
        if pd.isnull(df[trade_player].iloc[i]):
            df[trade_player].iloc[i] = '银行内部户交易'+df[trade_abstract[0]].iloc[i]

    trade_month = df[trade_time].dt.month
    trade_year = df[trade_time].dt.year
    df['交易年月'] = trade_year.astype(
        str)+'.'+trade_month.astype(str).str.zfill(2)

    return df



# 余额变化以及日均存款
def calavgres(dataframe):
    df = dataframe
    df[trade_time]
    df['date_parsed'] = df[trade_time].apply(lambda x: x.strftime('%Y-%m-%d'))
    df['date_parsed'] = pd.to_datetime(df['date_parsed'])
    
    day_range = pd.date_range(np.min(df['date_parsed']), np.max(df['date_parsed']))
    
    K_line = []
    X_ticks = []
    day_end = []
    
    end = df[left_amount].iloc[0]
    
    for date in day_range:
        day_trans = df[df['date_parsed'] == date]
        day_trans = day_trans.sort_values(by='date_parsed')
        if len(day_trans) != 0:
            if day_end!=[]:  
                begin = day_end[-1]
            else:
                begin = df[left_amount].iloc[0]
            end = day_trans[left_amount].iloc[-1]
            max_amount = np.max(day_trans[left_amount])
            min_amount = np.min(day_trans[left_amount])
            
        else:
            begin = end
            max_amount = end
            min_amount =end
            
        day_end.append(end)    
        each_day = [begin, end, min_amount,max_amount]
        each_day = np.array(each_day)/10000
        each_day = np.around(each_day, 2)
        K_line.append(each_day.tolist())
        X_ticks.append(str(date)[:10])
        
    day_end = np.around(np.array(day_end)/10000, 2)
    avg_res = np.around(np.mean(day_end),2)
    
    # 资金富余程度分数
    ## 日均资产绝对值 8分
    rich_score1 =  np.min([avg_res/30,8])
    
    ## 在日均资产平均值以上的天数比 3分
    rich_score2 = len(day_end[day_end>avg_res])/len(day_end)*3
    
    ## 在日均资产平均值50%以上的天数比 3分
    rich_score3 = len(day_end[day_end>(avg_res*0.5)])/len(day_end)*3
    
    ## 日末资产大于10w元天数 6分
    rich_score4 = len(day_end[day_end>10])/len(day_end)*6
    
    rich_score = rich_score1+rich_score2+rich_score3+rich_score4
    

    c = (
        Kline()
        .add_xaxis(X_ticks)
        .add_yaxis("每日余额变化情况", K_line)
        .set_global_opts(
            xaxis_opts=opts.AxisOpts(is_scale=True),
            yaxis_opts=opts.AxisOpts(name='万元',
                is_scale=True,
                splitarea_opts=opts.SplitAreaOpts(
                    is_show=True, areastyle_opts=opts.AreaStyleOpts(opacity=1)
                ),
            ),
            datazoom_opts=[opts.DataZoomOpts(range_start=0, range_end=100)],
            title_opts=opts.TitleOpts(
                title="Kline-每日余额变化情况", subtitle='日均资产余额：'+str(avg_res)+'万元', pos_left="right"),
        )
    )
    # c.render_notebook()

    line1 = (
        Line()
        .set_global_opts(
                        tooltip_opts=opts.TooltipOpts(trigger="axis"),
                        axispointer_opts=opts.AxisPointerOpts(
                            is_show=True, link=[{"xAxisIndex": "all"}]),
                        xaxis_opts=opts.AxisOpts(
                            type_="category", boundary_gap=False, axisline_opts=opts.AxisLineOpts(is_on_zero=True)),
                        yaxis_opts=opts.AxisOpts(name="万元"),)
        .add_xaxis(xaxis_data=X_ticks)
        .add_yaxis(
            series_name="每日终账户余额",
            y_axis=day_end,
            symbol_size=8,
            is_hover_animation=False,
            label_opts=opts.LabelOpts(is_show=False),
            linestyle_opts=opts.LineStyleOpts(width=2),
        )
    )
    
    c = c.overlap(line1)
    return c,rich_score


## 筛选关联公司及关联交易
def related_trade(dataframe):
    df = dataframe

    df_in = df[df[trade_in]>0]
    df_out = df[df[trade_out]>0]


    buyer_company = df_in.groupby(trade_player).sum()[trade_in]
    seller_company = df_out.groupby(trade_player).sum()[trade_out]

    related_company = set(buyer_company.index) & set(seller_company.index)
    related_trade = df[df[trade_player].isin(related_company)]


    related_trade_in = related_trade[related_trade[trade_in] > 0]
    related_trade_out = related_trade[related_trade[trade_out] > 0]

    related_trade_out_sum = related_trade_out.groupby(trade_player).sum()[trade_out].sort_index()
    related_trade_in_sum = related_trade_in.groupby(trade_player).sum()[trade_in].sort_index()

    # 进出资金比例大于阈值才筛选为关联交易对手
    threshold = 0.15
    up_threshold = (1-threshold)/threshold
    down_threshold = threshold/(1-threshold)
    ratio = related_trade_out_sum.values/related_trade_in_sum.values
    real_related_index = np.where((ratio<up_threshold) & (ratio>down_threshold))

    related_company = related_trade_out_sum.index[real_related_index].tolist()

    related_trade_in = related_trade_in[related_trade_in[trade_player].isin(related_company)]
    related_trade_out = related_trade_out[related_trade_out[trade_player].isin(related_company)]
    related_trade_out_sum = related_trade_out.groupby(trade_player).sum()[trade_out].sort_index()
    related_trade_in_sum = related_trade_in.groupby(trade_player).sum()[trade_in].sort_index()

    # 根据交易金额之和的绝对值筛选关联交易对手
    amount_threshold = 30 #万元
    sum_amount = (related_trade_out_sum.values+related_trade_in_sum.values)/10000
    real_related_index = np.where(sum_amount>amount_threshold)
    related_company = related_trade_out_sum.index[real_related_index].tolist()
    print(related_company)
    related_trade_in = related_trade_in[related_trade_in[trade_player].isin(related_company)]
    related_trade_out = related_trade_out[related_trade_out[trade_player].isin(related_company)]
    related_trade_out = related_trade_out.groupby(trade_player).sum()[trade_out].sort_index()
    related_trade_in = related_trade_in.groupby(trade_player).sum()[
        trade_in].sort_index()
    
    # 关联交易可视化
    labels = related_trade_in.index
    trade_in_values = related_trade_in.values/10000
    trade_out_values = related_trade_out.values/10000
    sum_values = trade_in_values+trade_out_values
    temp_df = pd.DataFrame(
        {"labels": labels, 'related_trade_in': trade_in_values,'related_trade_out':trade_out_values,'sum_values':sum_values})
    temp_df = temp_df.sort_values(by=['sum_values'],ascending=False)
    related_trade_ratio = (trade_out_values.sum()+trade_in_values.sum()) / \
        ((df[trade_in].sum()+df[trade_out].sum())/10000)
    related_trade_ratio = np.around(related_trade_ratio*100,2)

    related_bar = (
        Bar()
        .add_xaxis(temp_df.labels.values.tolist())
        .add_yaxis("关联交易入账", np.around(temp_df.related_trade_in.values.tolist(),2).tolist(), stack="stack1")
        .add_yaxis("关联交易出账", np.around(temp_df.related_trade_out.values.tolist(), 2).tolist(), stack="stack1")
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(title_opts=opts.TitleOpts(title="关联交易对手及金额", pos_top="1%",
                                                subtitle='关联交易流水占总金额比例：'+str(related_trade_ratio)+'%\n'+
                                                '关联交易入账:'+str(np.around(trade_in_values.sum(), 2))+'万元\n'+
                                                '关联交易出账:'+str(np.around(trade_out_values.sum(),2))+'万元'),
                        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=45, interval=0)),
                        yaxis_opts=opts.AxisOpts(name="交易金额：万元"),
                        datazoom_opts=opts.DataZoomOpts(range_start=0, range_end=100),)
    )
    grid_related = Grid(init_opts=opts.InitOpts(width="1200px", height="640px"))
    grid_related.add(related_bar, grid_opts=opts.GridOpts(pos_bottom="30%", pos_right='5%',
                                            pos_left='5%',pos_top='20%'), is_control_axis_index=True)
    return grid_related, {'related_company': related_company, 'buyer_company': buyer_company,
                          'seller_company': seller_company,'trade_in_values':trade_in_values,
                          'trade_out_values':trade_out_values}



# 真实上下游交易筛选归类
# 真实上下游交易筛选归类
def true_trade(dataframe,company_group):   
    df = dataframe
    buyer_company = company_group['buyer_company']
    seller_company = company_group['seller_company']
    related_company = company_group['related_company']
    
    all_company = set(buyer_company.index) | set(seller_company.index)
    normal_company = set(all_company) - set(related_company)
    normal_trade = df[df[trade_player].isin(normal_company)]

    normal_trade_in = normal_trade[normal_trade[trade_in] > 0]
    normal_trade_out = normal_trade[normal_trade[trade_out] > 0]

    normal_trade_out_sum = normal_trade_out.groupby(
        trade_player).sum()[trade_out].sort_values(ascending=False)
    normal_trade_in_sum = normal_trade_in.groupby(
        trade_player).sum()[trade_in].sort_values(ascending=False)

    in_count = normal_trade_in[trade_player].value_counts().to_frame()
    out_count = normal_trade_out[trade_player].value_counts().to_frame()
    in_trade = normal_trade_in_sum.to_frame()
    out_trade = normal_trade_out_sum.to_frame()

    in_trade = in_trade.join(in_count, on=trade_player, how='outer')
    in_trade = in_trade.rename(columns={trade_player: '交易次数'})
    out_trade = out_trade.join(out_count, on=trade_player, how='outer')
    out_trade = out_trade.rename(columns={trade_player: '交易次数'})


    
    in_step_freq = []
    for i in range(len(in_trade)):
        temp_df = df[trade_time].loc[df[trade_player] == in_trade.index[i]].values
        if len(temp_df) > 1:
            avg_step = []
            for j in range(len(temp_df)-1):
                time_step = temp_df[j+1] - temp_df[j]
                time_step = time_step.astype('timedelta64[D]')
                time_step = time_step / np.timedelta64(1, 'D')
                avg_step.append(time_step)
            avg_step = np.average(np.array(avg_step))

        else:
            avg_step = 0
        in_step_freq.append(np.around(avg_step, 2))
    in_trade['平均交易间隔时间'] = in_step_freq

    out_step_freq = []
    for i in range(len(out_trade)):
        temp_df = df[trade_time].loc[df[trade_player] == out_trade.index[i]].values
        if len(temp_df) > 1:
            avg_step = [] 
            for j in range(len(temp_df)-1):
                time_step = temp_df[j+1]- temp_df[j]
                time_step = temp_df[j+1] - temp_df[j]
                time_step = time_step.astype('timedelta64[D]')
                time_step = time_step / np.timedelta64(1, 'D')
                avg_step.append(time_step)
            avg_step = np.average(np.array(avg_step))
        else:
            avg_step = 0
        out_step_freq.append(np.around(avg_step, 2))
    out_trade['平均交易间隔时间'] = out_step_freq

    
    return normal_trade_in_sum,normal_trade_out_sum,in_trade,out_trade,normal_company

## 交易可视化
## 交易可视化
def trade_visual(normal_trade_in_sum,normal_trade_out_sum,in_trade,out_trade,company_group,top_visual=15):
    trade_in_values = company_group['trade_in_values']
    trade_out_values = company_group['trade_out_values']
    
    top_in_ration = (normal_trade_in_sum.values)[
        :top_visual].sum()/normal_trade_in_sum.values.sum()
    top_out_ration = (normal_trade_out_sum.values)[
        :top_visual].sum()/normal_trade_out_sum.values.sum()
    top_in_ration = np.around(top_in_ration*100, 2)
    top_out_ration = np.around(top_out_ration*100, 2)
    
    trade_sum = np.array([normal_trade_in_sum.values.sum(),
                    normal_trade_out_sum.values.sum()])
    
    
    colors = ['rgba(41, 52, 98,0.8)', 'rgba(242, 76, 76,0.8)', "rgb(24, 116, 152)"]

    x_data = in_trade.index.values.tolist()
    legend_list = ["交易金额", "平均交易间隔时间", "交易次数"]

    trade_amount = np.around((in_trade[trade_in].values/10000), 2).tolist()
    trade_step = in_trade['平均交易间隔时间'].values.tolist()
    trade_times = (in_trade['交易次数']).values.tolist()

    bar = (
        Bar(init_opts=opts.InitOpts(width="1200px", height="640px"))
        .set_global_opts(xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=45, interval=0)))
        .add_xaxis(x_data)
        .add_yaxis(
            series_name=legend_list[0], y_axis=trade_amount, yaxis_index=1, color=colors[0], gap="0%"
        )
        .add_yaxis(series_name=legend_list[1], y_axis=trade_step, yaxis_index=0, color=colors[1], gap="0%")
        .extend_axis(
            yaxis=opts.AxisOpts(
                name=legend_list[0],
                type_="value",
                position="right",
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color=colors[1])
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} 万元"),
            )
        )
        .extend_axis(
            yaxis=opts.AxisOpts(
                type_="value",
                name=legend_list[2],
                position="left",
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color=colors[2])
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} 次"),
                splitline_opts=opts.SplitLineOpts(
                    is_show=True, linestyle_opts=opts.LineStyleOpts(opacity=1)
                ),
            )
        )
        .set_global_opts(
            yaxis_opts=opts.AxisOpts(
                type_="value",
                name=legend_list[1],
                position="right",
                offset=80,
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color=colors[0])
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} 天"),

            ),
            tooltip_opts=opts.TooltipOpts(
                trigger="axis", axis_pointer_type="cross"),
            title_opts=opts.TitleOpts(title="下游客户交易情况",
                                    subtitle='下游入账总金额：'+str(np.around(trade_sum[0]/10000, 2))+'万元；总共交易对手数量：'+str(len(in_trade))+'；Top'+str(
                                        top_visual)+'交易对手金额占比：'+str(top_in_ration)+'%'),
            datazoom_opts=opts.DataZoomOpts(range_start=0, range_end=15))
    )

    line = (
        Line(init_opts=opts.InitOpts(width="1200px", height="640px"))
        .add_xaxis(xaxis_data=x_data)
        .add_yaxis(
            series_name=legend_list[2], y_axis=trade_times, yaxis_index=2, color=colors[2],
            linestyle_opts=opts.LineStyleOpts(width=2, color=colors[2]),
            itemstyle_opts=opts.ItemStyleOpts(color=colors[2])
        )
    )

    overlap = bar.overlap(line)
    grid = Grid(init_opts=opts.InitOpts(width="1200px", height="640px"))
    grid.add(overlap, grid_opts=opts.GridOpts(pos_bottom="30%", pos_right='20%',
                                            pos_left='10%', pos_top="15%"), is_control_axis_index=True)

    # 上游客户可视化
    colors = ['rgba(41, 52, 98,0.8)', 'rgba(242, 76, 76,0.8)', "rgb(24, 116, 152)"]

    x_data = normal_trade_out_sum.index.values.tolist()
    legend_list = ["交易金额", "平均交易间隔时间", "交易次数"]

    trade_amount = np.around((normal_trade_out_sum.values/10000), 2).tolist()
    trade_step = out_trade['平均交易间隔时间'].values.tolist()
    trade_times = (out_trade['交易次数']).values.tolist()

    bar = (
        Bar(init_opts=opts.InitOpts(width="1200px", height="640px"))
        .set_global_opts(xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=45, interval=0)))
        .add_xaxis(x_data)
        .add_yaxis(
            series_name=legend_list[1],
            y_axis=trade_step,
            yaxis_index=0,
            color=colors[1],
            gap="0%"
        )
        .add_yaxis(
            series_name=legend_list[0], y_axis=trade_amount, yaxis_index=1, color=colors[0], gap="0%"
        )
        .extend_axis(
            yaxis=opts.AxisOpts(
                name=legend_list[0],
                type_="value",
                position="right",
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color=colors[1])
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} 万元"),
            )
        )
        .extend_axis(
            yaxis=opts.AxisOpts(
                type_="value",
                name="交易次数",
                position="left",
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color=colors[2])
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} 次"),
                splitline_opts=opts.SplitLineOpts(
                    is_show=True, linestyle_opts=opts.LineStyleOpts(opacity=1)
                ),
            )
        )
        .set_global_opts(
            yaxis_opts=opts.AxisOpts(
                type_="value",
                name=legend_list[1],
                position="right",
                offset=80,
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color=colors[0])
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} 天"),

            ),
            tooltip_opts=opts.TooltipOpts(
                trigger="axis", axis_pointer_type="cross"),
            title_opts=opts.TitleOpts(title="上游客户交易情况",
                                    subtitle='上游出账总金额：'+str(np.around(trade_sum[1]/10000, 2))+'万元；总共交易对手数量：'+str(len(out_trade))+'；Top'+str(
                                        top_visual)+'交易对手金额占比：'+str(top_out_ration)+'%'),
            datazoom_opts=opts.DataZoomOpts(range_start=0, range_end=15))
    )

    line = (
        Line(init_opts=opts.InitOpts(width="1200px", height="640px"))
        .add_xaxis(xaxis_data=x_data)
        .add_yaxis(
            series_name=legend_list[2], y_axis=trade_times, yaxis_index=2, color=colors[2],
            linestyle_opts=opts.LineStyleOpts(width=2, color=colors[2]),
            itemstyle_opts=opts.ItemStyleOpts(color=colors[2])
        )
    )

    overlap1 = bar.overlap(line)
    grid1 = Grid(init_opts=opts.InitOpts(width="1200px", height="640px"))
    grid1.add(overlap1, grid_opts=opts.GridOpts(pos_bottom="30%", pos_right='20%',
                                            pos_left='10%', pos_top="15%"), is_control_axis_index=True)

    pie_ratio = (
        Pie()
        .add("", [['关联交易入账', np.around(trade_in_values.sum(), 2)], ['关联交易出账', np.around(trade_out_values.sum(), 2)],
                ['正常交易入账', np.around(trade_sum[0]/10000, 2)], ['正常交易出账', np.around(trade_sum[1]/10000, 2)]])
        .set_global_opts(title_opts=opts.TitleOpts(title="出入账比例"))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}万元\n({d}%)"))
    )
    
    ## 上下游交易金额大于1万的对手数量分数
    normal_trade_in_sum_1w = normal_trade_in_sum[normal_trade_in_sum>10000.0]
    normal_trade_out_sum_1w = normal_trade_out_sum[normal_trade_out_sum>10000.0]
    s_upstream = np.sqrt(np.min([len(normal_trade_in_sum_1w),100])) *1# 10分
    s_downstream = np.sqrt(np.min([len(normal_trade_in_sum_1w),100])) *0.4 # 4分

    # 交易集中程度 6分
    top5ratio = (1-(normal_trade_in_sum.values)[:5].sum()/normal_trade_in_sum.values.sum()) * 3.5 
    top10ratio = (1-(normal_trade_in_sum.values)[:10].sum()/normal_trade_in_sum.values.sum()) * 2
    top15ratio = (1-(normal_trade_in_sum.values)[:15].sum()/normal_trade_in_sum.values.sum()) * 0.5
    
    # 交易对手分数 20分
    print(s_upstream,s_downstream,top5ratio,top10ratio,top15ratio)
    rival_score = s_upstream+s_downstream+top15ratio+top5ratio+top10ratio
    
    # 关联交易程度 超过50%为0分
    in_score = np.min(1-(trade_in_values.sum()/(trade_in_values.sum()+trade_sum[0]/10000))*2,0)
    out_score = np.min(1-((trade_out_values.sum()/(trade_sum[1]/10000+trade_out_values.sum())))*2,0)
    related_score = in_score*0.8 + out_score*0.2
    related_score = np.max([related_score,0])*20
    
    return grid,grid1,pie_ratio,rival_score,related_score


def four_indicator_compute(df):
    # 水费
    water_fee = df[df['备注'].str.contains('水费') & df['借方金额(支出)'] != 0]
    water_fee = water_fee.groupby('交易年月')['借方金额(支出)'].sum()

    # 电费
    electric_fee = df[df['备注'].str.contains('电费') & df['借方金额(支出)'] != 0]
    electric_fee = electric_fee.groupby('交易年月')['借方金额(支出)'].sum()

    # 税
    tax = df[df['备注'].str.contains('税') & df['借方金额(支出)'] != 0]
    tax = tax.groupby('交易年月')['借方金额(支出)'].sum()

    # 工资
    salary = df[(df['备注'].str.contains('工资') |
                 df['备注'].str.contains('代发')) & df['借方金额(支出)'] != 0]
    salary = salary.groupby('交易年月')['借方金额(支出)'].sum()

    # 五险一金
    insurance = df[(df['备注'].str.contains('保险') |
                    df['备注'].str.contains('公积金')) & df['借方金额(支出)'] != 0]
    insurance = insurance.groupby('交易年月')['借方金额(支出)'].sum()

    indicators = [water_fee, electric_fee, tax, salary, insurance]

    # 补充空缺
    time_step = pd.date_range(start=min(df['交易年月']), end=max(df['交易年月']),
                              freq='M',).strftime("%Y.%m").to_list()
    for i in range(len(time_step)):
        for j in range(len(indicators)):
            if time_step[i] not in indicators[j].index:
                indicators[j][time_step[i]] = 0

        indicators[j] = indicators[j].sort_index()

    four_indicator_bar = (
        Bar()
        .add_xaxis(time_step)
        .add_yaxis('水费', np.around(indicators[0].values/10000, 2).tolist(), stack="stack1",
                   label_opts=opts.LabelOpts(position='top'))
        .set_series_opts()

        .add_yaxis("电费", np.around((indicators[1])/10000, 2).tolist(), stack="stack2",
                   label_opts=opts.LabelOpts(position='top'))
        .add_yaxis("税", np.around((indicators[2])/10000, 2).tolist(), stack="stack3",
                   label_opts=opts.LabelOpts(position='top'))
        .add_yaxis("工资", np.around((indicators[3])/10000, 2).tolist(), stack="stack4",
                   label_opts=opts.LabelOpts(position='top'))
        .add_yaxis("五险一金", np.around((indicators[4])/10000, 2).tolist(), stack="stack4",

                   label_opts=opts.LabelOpts(position='top'))
        .set_global_opts(title_opts=opts.TitleOpts(title="中小微企业四项指标统计"),
                         yaxis_opts=opts.AxisOpts(name='万元'),
                         datazoom_opts=opts.DataZoomOpts(
            range_start=0, range_end=100),
        )

    )
    return(four_indicator_bar, indicators)


#每月交易回笼统计
def statbymonth(dataframe, normal_company, output_excel=True):
    df = dataframe
    holdshare = []
    company_length = 4
    month_group = list(df.groupby(['交易年月']))
    company_trade = df[df[trade_player].str.len() > company_length]
    company_trade = company_trade[company_trade[trade_player].isin(
        normal_company)]
    company_trade = company_trade[~company_trade[trade_player].isin(
        holdshare)]
    stat_in_month = company_trade.groupby(['交易年月']).sum()[trade_in]
    stat_out_month = company_trade.groupby(['交易年月']).sum()[trade_out]

    time_step = pd.date_range(start=min(df['交易年月']), end=max(df['交易年月']),
                              freq='M').strftime("%Y.%m").to_list()

    for i in range(len(time_step)):
        if time_step[i] not in stat_in_month.index:
            stat_in_month[time_step[i]] = 0
            stat_out_month[time_step[i]] = 0
    stat_in_month = stat_in_month.sort_index()
    stat_out_month = stat_out_month.sort_index()

    #  计算中小微企业四项指标
    four_indicator_bar, indicator = four_indicator_compute(df)

    if output_excel == True:
        stat_excel_path = './output/销售回笼统计.xlsx'
        if os.path.exists(stat_excel_path):
            os.remove(stat_excel_path)
        writer = pd.ExcelWriter(stat_excel_path)
        print('统计月份:', len(month_group), len(stat_in_month))
        for i in range(len(month_group)):
            month_group[i][1].to_excel(writer, month_group[i][0])
        # month_group[0][1]
        writer.save()
        wb = xw.Book(stat_excel_path)
        for i in range(len(month_group)):
            sheet = wb.sheets[month_group[i][0]]
            last_col = sheet.used_range.shape[0]
            sheet.autofit(axis="columns")
            for j in range(2, sheet.used_range.shape[0]):
                if sheet.range((j, 6)).value == None:
                    continue
                if len(sheet.range((j, 6)).value) > company_length:
                    if sheet.range((j, 6)).value in normal_company:
                        if sheet.range((j, 6)).value not in holdshare:
                            if sheet.range((j, 3)).value > 0:
                                sheet.range((j, 1), (j, 8)).color = (
                                    230, 184, 183)
            sheet.range((last_col+1, 3)).value = stat_in_month[i]
            sheet.range((last_col+1, 2)).value = '销售回笼收入合计'
        try:
            wb.sheets.add('每月销售收入台账')
        except:
            pass
        sheet = wb.sheets['每月销售收入台账']
        sheet.range('A1').value = '交易年月'
        sheet.range('A2').value = '当月累计销售回笼金额'
        sheet.range('B1').value = stat_in_month.index.values
        sheet.range('B2').value = stat_in_month.values

        sheet.range('A3').value = '水费'
        sheet.range('B3').value = indicator[0].values
        sheet.range('A4').value = '电费'
        sheet.range('B4').value = indicator[1].values
        sheet.range('A5').value = '税费'
        sheet.range('B5').value = indicator[2].values
        sheet.range('A6').value = '工资'
        sheet.range('B6').value = indicator[3].values
        sheet.range('A7').value = '五险一金'
        sheet.range('B7').value = indicator[4].values

        last_row = sheet.used_range.shape[1]
        sheet.range((1, last_row+1)).value = '总计'
        sheet.range((2, last_row+1)).value = stat_in_month.sum()
        sheet.autofit(axis="columns")
        wb.save(stat_excel_path)
        wb.close()

    in_out_bar = (
        Bar()
        .add_xaxis(stat_in_month.index.values.tolist())
        .add_yaxis('销售回笼金额', np.around(stat_in_month.values/10000, 2).tolist(), stack="stack1",
                   label_opts=opts.LabelOpts(position='top'))
        .set_series_opts()
        .add_yaxis("上游支付金额", np.around((0-stat_out_month.values)/10000, 2).tolist(), stack="stack1",
                   label_opts=opts.LabelOpts(position='bottom'))
        # .set_series_opts(label_opts=opts.LabelOpts(position='down'))
        .set_global_opts(title_opts=opts.TitleOpts(title="每月进出账统计"),
                         )
    )

    # 回款平稳性分数
    ## 标准差等级分数 3分
    std_stat_in_month = stat_in_month/np.linalg.norm(stat_in_month)
    stable1 = np.min([(1-np.mean(np.std(std_stat_in_month)))*5, 5])

    ## 回款金额在50%平均值-150%平均值之间的月份所占比例 10分
    month_count = stat_in_month[(stat_in_month > 0.5*np.mean(stat_in_month))
                                & (stat_in_month < 1.5*np.mean(stat_in_month))].count()
    stable2 = month_count/len(stat_in_month)*10

    ## 最大回款与最小回款比值，3分扣减
    stable3 = 3 - np.min([(np.max(stat_in_month) -
                           np.min(stat_in_month))/(np.min(stat_in_month)+1), 5])

    ## 每个月和平均值的差距 4分扣减
    stable4 = 4 - np.min([np.mean(np.abs(stat_in_month -
                                         np.mean(stat_in_month))/np.mean(stat_in_month)), 4])

    ### 平稳性分数最终值
    stable_score = stable1+stable2+stable3+stable4

    # 经营回笼趋势
    # 以3个月为窗口，滑动计算较上个周期的增长值 4分
    all_compare_ratio = []
    for i in range(len(stat_in_month)-4):
        compare_ratio = np.min([(np.sum(stat_in_month[i+1:i+4])-np.sum(
            stat_in_month[i:i+3]))/(np.sum(stat_in_month[i:i+3])+1), 1])
        compare_ratio = np.max([compare_ratio, -1])
        all_compare_ratio.append(compare_ratio)
    trend_score1 = (np.mean(all_compare_ratio)+1)/2*4

    # 全阶段曲线拟合 a值 3分
    x = np.arange(0, len(stat_in_month), 1)
    a, b = np.polyfit(x, stat_in_month, deg=1)
    trend_score2 = np.min([np.max([a+1, 0])*2, 5])

    # 纳税情况拟合 3分
    tax = indicator[2].values
    x = np.arange(0, len(tax), 1)
    a, b = np.polyfit(x, tax, deg=1)
    trend_score3 = np.min([np.max([a+1, 0]), 3])

    # 四项指标缴纳月数 5分
    trend_score4 = 0
    for i in range(len(indicator)):
        temp = indicator[i]
        ratio = indicator[i][indicator[i] > 0].count()/len(indicator[i])
        trend_score4 += np.min([ratio, 1])

    # 五险一金缴纳比例 2分
    ratio = indicator[4]/(indicator[3]+1)
    ratio = ratio[ratio < 1]
    trend_score5 = np.min([np.mean(ratio)*20, 2])

    # 近6个月回笼趋势 3分
    if len(stat_in_month) >= 6:
        x = np.arange(0, 6, 1)
        a, b = np.polyfit(x, stat_in_month[-6:], deg=1)
        trend_score6 = np.min([np.max([a+1, 0])*2, 3])
    else:
        trend_score6 = trend_score2

    print('trend_score:', trend_score1, trend_score2,
          trend_score3, trend_score4, trend_score5, trend_score6)

    trend_score = trend_score1+trend_score2 + \
        trend_score3+trend_score4+trend_score5+trend_score6

    return in_out_bar, four_indicator_bar, stable_score, trend_score, stat_in_month



def wordcloud(df):
    text_summary = ','.join(df['备注'])
    jieba.load_userdict('./dictionary.txt')

    cut = jieba.analyse.extract_tags(text_summary,topK=50,withWeight=True)
    '''从文本中生成词云图'''
    wordcloud = WD(background_color='white',  # 背景色为白色
                height=400,  # 高度设置为400
                width=800,  # 宽度设置为800
                scale=20,  # 长宽拉伸程度设置为20
                font_path="/System/Library/Fonts/STHeiti Medium.ttc",
                repeat=False,
                relative_scaling=0.2).generate(text_summary)


    data = [(key,float(wordcloud.words_[key])) for key in wordcloud.words_]
    data.extend(cut)
    cloud_html = (
        WordCloud(init_opts=opts.InitOpts(width="1200px", height="640px"))
        .add(series_name="热点分析", data_pair=data,word_size_range=[10,150])
        .set_global_opts(
            title_opts=opts.TitleOpts(
                title="交易摘要词云图", title_textstyle_opts=opts.TextStyleOpts(font_size=23)
            ),
            tooltip_opts=opts.TooltipOpts(is_show=True),
        )
    )
    return cloud_html



#雷达图


def plot_radar(scores):
    radar_plot = (
        Radar()
        .add_schema(
            schema=[
                opts.RadarIndicatorItem(name="交易对手多样性", max_=20),
                opts.RadarIndicatorItem(name="回款稳定性", max_=20),
                opts.RadarIndicatorItem(name="关联交易情况", max_=20),
                opts.RadarIndicatorItem(name="经营回笼趋势", max_=20),
                opts.RadarIndicatorItem(name="资金富余程度", max_=20),

            ],
            splitarea_opt=opts.SplitAreaOpts(
                is_show=True, areastyle_opts=opts.AreaStyleOpts(opacity=1))
        )
        .add("流水各项情况评分", [scores], areastyle_opts=opts.AreaStyleOpts(opacity=0.1))
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))

        .set_global_opts(
            title_opts=opts.TitleOpts(
                title="流水评分雷达图", subtitle='总体评分：'+str(np.around(np.sum(scores)))),
        )
    )
    return radar_plot

# 全流程


def workflow(dataframe,output_excel=False):
    df = dataframe
    line1,trend_score = calavgres(df)
    grid_related,company_group = related_trade(df)
    normal_trade_in_sum,normal_trade_out_sum,in_trade,out_trade,normal_company = true_trade(df,company_group)
    grid,grid1,pie_ratio,rival_score,related_score = trade_visual(normal_trade_in_sum,normal_trade_out_sum,
                                        in_trade,out_trade,company_group,top_visual=15)
    in_out_bar, four_indicator_bar, stable_score, compare_score, stat_in_month = statbymonth(
        df, normal_company, output_excel=output_excel)

    SCORE1 = rival_score
    SCORE2 = stable_score
    SCORE3 = related_score
    SCORE4 = compare_score
    SCORE5 = trend_score

    scores = np.around(np.array([SCORE1,SCORE2,SCORE3,SCORE4,SCORE5]),2).tolist()
    print(scores)
    radar_plot = plot_radar(scores)
    wordcloud_html = wordcloud(df)


    tab = Tab('流水分析情况')
    tab.add(radar_plot,'流水整体评分情况')
    tab.add(line1, '账户余额走势')
    tab.add(grid_related, '关联交易情况')
    tab.add(pie_ratio, '出入账比例')
    tab.add(in_out_bar, '每月销售回笼情况统计')
    tab.add(four_indicator_bar, '中小微企业四项指标统计')
    tab.add(grid, '下游交易情况')
    tab.add(grid1, '上游交易情况')
    tab.add(wordcloud_html,'交易摘要词云图')
    tab.render('./templates/流水可视化分析.html')
    return tab

