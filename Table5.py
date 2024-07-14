#20240707完成通过测试

import pandas as pd
import os
# path=r"F:\pandas\Oracle\check0604\output110605.csv"
pd.set_option('display.float_format', '{:,.2f}'.format)
# path=r"F:\pandas\Oracle\LASTCAR0531\222324123417000601csv.csv"
outpath=r"F:\pandas\Oracle\CARAUTOREPORT\222324_12_12345险别险种\output.xlsx"
pathXL=r"F:\pandas\Oracle\CARAUTOREPORT\222324_12_12345险别险种\222324_12_12345new060.csv" #险类文件路径，22，23年全年数据，24年1月-5月数据 
pathXB=r"F:\pandas\Oracle\CARAUTOREPORT\222324_12_12345险别险种\merged_tableF.csv"#险别文件路径，22，23年全年数据，24年1月-5月数据，与险别码表merge并且剔除保费为空的行，15013

df = pd.DataFrame()
df.to_excel(outpath)
columnsXB = [
    "保单号","险种名称","险别名称","保费不含税","险种代码","险别代码","险种名称一",
    "CLOSEDDATE", "MAPPINGPROVINCENAME", "MAPPINGCITYNAME", "CLASSCODE",
    "CLASSNAME", "RISKCODE", "RISKNAME", "ENDORSETIMES",
    "ENDORSENO", "CANCELTYPE", "CANCELNAME", "ENDORTYPE", "ENDORNAME",
    "STARTDATE", "USENATURECODE", "CARKINDCODE", "CARTYPE","SIGNPREMIUM","FRAMENO",
]
columnsXL = [
    
    "CLOSEDDATE", "MAPPINGPROVINCENAME", "MAPPINGCITYNAME", "CLASSCODE",
    "CLASSNAME", "RISKCODE", "RISKNAME", "ENDORSETIMES",
    "ENDORSENO", "CANCELTYPE", "CANCELNAME", "ENDORTYPE", "ENDORNAME",
    "STARTDATE", "USENATURECODE", "CARKINDCODE", "CARTYPE","SIGNPREMIUM","FRAMENO",
    
    'TONCOUNTMENT',
    'USEYEARSMENT',
    'PURCHASEPRICEMENT',
    'JQNCD',
    'SYNCD',
    
    'EARNEDPREMIUMS',
    'SUMPAYINDIRFEE',
   'OUTSTANDINGPAYINDIRFEE',
    
    
    
    
]
# 'SEATCOUNTMENT':"座位数分段",
#                     'TONCOUNTMENT':"吨位数分段",
#                     'USEYEARSMENT':"使用年限分段",
#                     'PURCHASEPRICEMENT':"新车购置价分段",
# 'JQNCD':"交强险NCD",
#                     'SYNCD':"商业险NCD",
# 读取CSV文件
dfXL = pd.read_csv(pathXL,encoding='gbk', usecols=columnsXL)
print(dfXL.info())

dfXB = pd.read_csv(pathXB,encoding='gbk', usecols=columnsXB)
print(dfXB.info())
dfXB.rename(columns={'保费不含税': '签单保费'}, inplace=True)
#----------------------------------------------------------------------------------------------------------DATA CHECK
#注意 险类和险别总签单保费数相等
total_signpremiumXL = dfXL['SIGNPREMIUM'].sum()#签单保费（不含税）总和,险类文件数据校验
print(total_signpremiumXL)#险类总签单保费不含税吻合 1250085061.28

total_signpremiumXB = dfXB['签单保费'].sum()#签单保费（不含税）总和,险别文件数据校验
print(total_signpremiumXB)#险类总签单保费不含税吻合 1250085061.28
#----------------------------------------------------------------------------------------------------------DATA CHECK
#------------------------------------------------------------------------------------------------------------------------------------------------------险类数据预处理
def classify_insurance(riskname):
    if '强制' in riskname:
        return '交强险'
    elif '商业' in riskname:
        return '商业险'
    else:
        return '其他'

dfXL['险种'] = dfXL['RISKNAME'].apply(classify_insurance)#“险种”列 分为1.交强险 2.商业险

dfXL['STARTDATE'] = pd.to_datetime(dfXL['STARTDATE'], format='mixed', dayfirst=True)
dfXB['STARTDATE'] = pd.to_datetime(dfXB['STARTDATE'], format='mixed', dayfirst=True)
# 将 STARTDATE 设置为索引
dfXL.set_index('STARTDATE', inplace=True)
dfXB.set_index('STARTDATE', inplace=True)
# 提取年份和月份
dfXL['年份'] = dfXL.index.year
dfXL['月份'] = dfXL.index.month

dfXB['年份'] = dfXB.index.year
dfXB['月份'] = dfXB.index.month



#---------------------------------------配置车种分类，车种分类字段数据库没得，得手动与CD_CARKIND匹配⬇
dfXL['combined_col'] = dfXL.apply(lambda row: f"{row['USENATURECODE']}_{row['CARKINDCODE']}", axis=1)
dfXB['combined_col'] = dfXB.apply(lambda row: f"{row['USENATURECODE']}_{row['CARKINDCODE']}", axis=1)
print(dfXL['combined_col'].unique())
print(dfXB['combined_col'].unique())
CarKindList = {
    '8D_A0': 'A家庭自用车', '8B_A0': 'B非营业客车', '8C_A0': 'B非营业客车', '9A_A0': 'C营业客车',
    '9B_A0': 'C营业客车', '9C_A0': 'C营业客车', '9K_A0': 'C营业客车', '98_H1': 'D非营业货车',
    '98_H0': 'D非营业货车', '98_G0': 'D非营业货车', '98_H3': 'D非营业货车', '98_H2': 'D非营业货车',
    '98_H4': 'D非营业货车', '98_H6': 'D非营业货车', '98_H5': 'D非营业货车', '97_H4': 'E营业货车',
    '97_H5': 'E营业货车', '97_H0': 'E营业货车', '97_H6': 'E营业货车', '97_H2': 'E营业货车',
    '97_H1': 'E营业货车', '97_H3': 'E营业货车', '97_G0': 'E营业货车', '97_TA': 'F特种车',
    '97_TB': 'F特种车', '97_T8': 'F特种车', '97_T9': 'F特种车', '97_TC': 'F特种车', '97_TF': 'F特种车',
    '97_TG': 'F特种车', '97_TD': 'F特种车', '97_TE': 'F特种车', '97_T7': 'F特种车', '97_T1': 'F特种车',
    '97_T5': 'F特种车', '97_T6': 'F特种车', '97_T2': 'F特种车', '97_T4': 'F特种车', '97_TT': 'F特种车',
    '97_TU': 'F特种车', '97_TR': 'F特种车', '97_TS': 'F特种车', '97_TV': 'F特种车', '97_TY': 'F特种车',
    '97_TZ': 'F特种车', '97_TW': 'F特种车', '97_TX': 'F特种车', '97_TQ': 'F特种车', '97_TJ': 'F特种车',
    '97_TK': 'F特种车', '97_TH': 'F特种车', '97_TI': 'F特种车', '97_TL': 'F特种车', '97_TO': 'F特种车',
    '97_TP': 'F特种车', '97_TM': 'F特种车', '97_TN': 'F特种车', '97_G3': 'F特种车', '97_G1': 'F特种车',
    '97_G2': 'F特种车', '97_G4': 'F特种车', '97_TQ1': 'F特种车', '97_TQ2': 'F特种车', '97_TQ3': 'F特种车',
    '98_TN': 'F特种车', '98_G4': 'F特种车', '98_TK': 'F特种车', '98_TJ': 'F特种车', '98_TM': 'F特种车',
    '98_TL': 'F特种车', '98_TO': 'F特种车', '98_TV': 'F特种车', '98_TU': 'F特种车', '98_TW': 'F特种车',
    '98_TZ': 'F特种车', '98_TY': 'F特种车', '98_TX': 'F特种车', '98_TT': 'F特种车', '98_TQ': 'F特种车',
    '98_TP': 'F特种车', '98_G3': 'F特种车', '98_G2': 'F特种车', '98_G1': 'F特种车', '98_TS': 'F特种车',
    '98_TR': 'F特种车', '98_TI': 'F特种车', '98_T4': 'F特种车', '98_T2': 'F特种车', '98_T6': 'F特种车',
    '98_T5': 'F特种车', '98_T1': 'F特种车', '98_T7': 'F特种车', '98_TE': 'F特种车', '98_TD': 'F特种车',
    '98_TF': 'F特种车', '98_TH': 'F特种车', '98_TG': 'F特种车', '98_T8': 'F特种车', '98_T9': 'F特种车',
    '98_TA': 'F特种车', '98_TC': 'F特种车', '98_TB': 'F特种车', '98_TQ1': 'F特种车', '98_TQ2': 'F特种车',
    '98_TQ3': 'F特种车', '97_N0': 'G其他', '97_SL': 'G其他', '97_M2': 'G其他', '97_M9': 'G其他',
    '97_J0': 'G其他', '97_J2': 'G其他', '97_J1': 'G其他', '97_J3': 'G其他', '97_M1': 'G其他',
    '97_M0': 'G其他', '98_J2': 'G其他', '98_J1': 'G其他', '98_J0': 'G其他', '98_M0': 'G其他',
    '98_M2': 'G其他', '98_M1': 'G其他', '98_M9': 'G其他', '98_SL': 'G其他', '98_J3': 'G其他',
    '98_N0': 'G其他','98_M75':'G其他'
}#====================================================此对照表版本较老，'98_M75':'G其他' 对应功率大于4Km小于8km的电动摩托车，于2024年6月6日手动添加
	
dfXL['车种分类'] = dfXL['combined_col'].map(CarKindList)
dfXB['车种分类'] = dfXB['combined_col'].map(CarKindList)#查看车种分类
# df4=df['车种分类'].unique()
# print(df4)
# df4=df['车种分类'].value_counts().sum()
# print(df4)
# missing_values = df[df['车种分类'].isna()]['combined_col'].unique()
# print("没有匹配到的值:", missing_values)
column_mapping = {
                    'CLOSEDDATE':"截止日期",
                    'MAPPINGPROVINCENAME':"省级业务归属区划",
                    'MAPPINGCITYNAME':"市级业务归属区划",
                    'MAPPINGTOWNNAME':"县区业务归属区划",
                    'CLASSCODE':"险类代码",
                    'CLASSNAME':"险类名称",
                    'RISKCODE':"险种代码",
                    'RISKNAME':"险种名称",
                    'COMCODE2':"业务归属二级机构代码",
                    'COMNAME2':"业务归属二级机构名称",
                    'COMCODE3':"业务归属三级机构代码",
                    'COMNAME3':"业务归属三级机构名称",
                    'COMCODE4':"业务归属四级机构代码",
                    'COMNAME4':"业务归属四级机构名称",
                    'COMCODE':"业务归属机构代码",
                    'COMCNAME':"业务归属机构名称",
                    'HANDLERCODE':"归属业务员代码",
                    'HANDLERNAME':"归属业务员名称",
                    'POLICYNO':"保单号",
                    'ENDORSETIMES':"批改次数",
                    'ENDORSENO':"批单号",
                    'CANCELTYPE':"是否注销",
                    'CANCELNAME':"是否注销名称",
                    'ENDORTYPE':"是否退保",
                    'ENDORNAME':"是否退保名称",
                    'BUSINESSTYPE':"涉农标识",
                    'BUSINESSNAME':"涉农标识名称",
                    'BUSINESSTYPE1':"政策/商业标志",
                    'BUSINESSNAME1':"政策/商业标志名称",
                    'GOVERNMENTFLAG':"政府合作项目",
                    'GOVERNMENTNAME':"政府合作项目名称",
                    'RENEWFLAG':"续保标识",
                    'RENEWNAME':"续保标识名称",
                    'ISSEEFEEFLAG':"见费出单标识",
                    'ISSEEFEENAME':"见费出单标识名称",
                    'COINSFLAG':"联共保标志",
                    'COINSNAME':"联共保标志名称",
                    'CHIEFFLAG':"是否首席",
                    'CHIEFNAME':"是否首席名称",
                    'OWNCOINSRATE':"我司份额",
                    'BUSINESSCATEGORY':"业务大类",
                    'RESOLUTION':"合同争议解决方式",
                    'RESOLUTIONNAME':"合同争议解决方式名称",
                    'MAINBUSINESSNATURE':"业务来源大类代码",
                    'MAINBUSINESSNATURENAME':"业务来源大类名称",
                    'BUSINESSNATURE':"业务来源小类代码",
                    'BUSINESSNATURENAME':"业务来源小类名称",
                    'AGENTCODE':"代理人代码",
                    'AGENTNAME':"代理人名称",
                    'AGENTAGREEMENTNO':"代理协议号",
                    'APPLICODE':"投保人客户代码",
                    'APPLINAME':"投保人客户名称",
                    'APPLITYPE':"投保人客户类型",
                    'APPLITYPENAME':"投保人客户类型名称",
                    'APPLIRISKTYPE':"投保人风险等级",
                    'INSUREDCODE':"被保险人客户代码",
                    'INSUREDNAME':"被保险人客户名称",
                    'INSUREDTYPE':"被保险人客户类型",
                    'INSUREDTYPENAME':"被保险人客户类型名称",
                    'INSUREDRISKTYPE':"被保险人风险等级",
                    'STARTDATE':"起保日期",
                    'ENDDATE':"终保日期",
                    'ORIUWENDDATE':"原始保单核保日期",
                    'ORIUWNAME':"原始保单核保人名称",
                    'ORISIGNDATE':"原始保单核单日期",
                    'ORIOPERATEDATE':"原始保单签单日期",
                    'LISENSENO':"车牌号码",
                    'FRAMENO':"车架号",
                    'ENGINENO':"发动机号",
                    'USENATURECODE':"车辆使用性质",
                    'CARKINDCODE':"车辆种类",
                    'CAROWNERNATURE':"车主性质",
                    'IDENTITYSUREDCAR':"被保险人与车辆关系",
                    'BRANDNATURE':"号牌种类",
                    'COUNTRYCODE':"国别性质",
                    'CARTYPECNAME':"交管车辆类型",
                    'MODELNAME':"车型名称",
                    'BRANDNAME':"品牌名称",
                    'CARSERIESNAME':"车系名称",
                    'CARTYPE':"车型种类",
                    'SEATCOUNT':"座位数",
                    'ENROLLDATE':"初登日期",
                    'USEYEARS':"使用年限",
                    'PURCHASEPRICE':"新车购置价",
                    'CAPACITYNUM':"核定载客量",
                    'QUALITYNUM2':"QUALITYNUM2",
                    'WHOLEWEIGHT':"整备质量",
                    'EXHAUSTSCALE':"排量/功率",
                    'ENERGYTYPE':"能源种类",
                    'SPECIALFLAG':"特殊车投保标志",
                    'MODELCODE':"车型代码",
                    'BUSINESSMODELCODE':"行业车型编码",
                    'CHGOWNERDATE':"过户日期",
                    'LOANVEHICLEFLAG':"车贷投保多年标志",
                    'CHGOWNERFLAG':"是否过户车",
                    'LOADCAR':"贷款车辆",
                    'LOCALCARFLAG':"外地车标志",
                    'NEWEQUIPMENTFLAG':"新增设备标志",
                    'APPNTVIPFLAG':"投保人VIP标志",
                    'INSUREDVIPFLAG':"被保险人VIP标志",
                    'COEFILLEGAL':"交通违法系数",
                    'LASTYEARCLAIMTIMES':"上年理赔次数",
                    'LASTYEARCLAIMSUMS':"上年理赔金额",
                    'NOPAYRATIO':"无赔优系数",
                    'AUTOCHANNELRATIO':"自主渠道系数",
                    'AUTOUWRATIO':"自主核保系数",
                    'AUTOPRODUCTRATIO':"自主系数乘积",
                    'SUGGESTEDDISCOUNT':"初始建议折扣",
                    'EXPECTEDDISCOUNTED':"期望折扣",
                    'DISCOUNT':"折扣",
                    'GROUPAGREEMENTNO':"团单协议号",
                    'CARAGREEMENTNO':"车队协议号",
                    'TAXSITUATION':"纳税情况",
                    'JQVALIDFALG':"交强险即时生效标志",
                    'VVTAXPAYTYPE':"车船税缴费类型",
                    'PAIDFREECERTIFICATE':"完税凭证号",
                    'RELIEFREASON':"减免原因",
                    'RELIEFPLAN':"减免方案",
                    'TAXPAYMENTDATE':"完税凭证填发日期",
                    'TAXPAYMENTREGIONCODE':"完税凭证地区代码",
                    'CARCHECKSTATUS':"验车情况",
                    'CARCHECKOPERTOR':"验车人",
                    'CARCHECKDATE':"验车日期",
                    'NOCHECKREASON':"免验原因",
                    'SUMCURRENCY':"汇总币别",
                    'SUMCURRENCYEXCHANGERATE':"汇总币别兑换率",
                    'SIGNAMOUNT':"签单保额",
                    'SIGNPREMIUM':"签单保费",
                    'TAX':"税额",
                    'SIGNPREMIUMOFTAX':"签单保费（含税）",
                    'SALVATIONFEEFLAG':"手续费标识",
                    'SALVATIONFEENAME':"手续费标识名称",
                    'DISRATE':"手续费比例",
                    'COMMCHARGE':"手续费金额",
                    'BASEPERFORMANCEFLAG':"绩效标识",
                    'BASEPERFORMANCENAME':"绩效标识名称",
                    'BASEPERFORMANCERATE':"绩效比例",
                    'BASEPERFORMANCE':"绩效金额",
                    'REGISTNUMBER':"立案件数",
                    'SUMPAY':"已决赔款",
                    'SUMPAYDIRECTLYFEE':"已决直接理赔费用",
                    'SUMPAYINDIRFEE':"已决赔款含直费",
                    'OUTSTANDINGPAY':"未决赔款",
                    'OUTSTANDINGDIRECTLYFEE':"未决直接理赔费用",
                    'OUTSTANDINGPAYINDIRFEE':"未决赔款含直费",
                    'CREATEDATE':"创建日期",
                    'UPDATEDATE':"修改日期",
                    'APPNTVIPNAME':"投保人vip标志名称",
                    'INSUREDVIPNAME':"被保险人vip标志名称",
                    'USENATURENAME':"车辆使用性质名称",
                    'CARKINDNAME':"车辆种类名称",
                    'LOANVEHICLENAME':"车贷投保多年标志名称",
                    'TONCOUNT':"吨位数",
                    'SEATCOUNTMENT':"座位数分段",
                    'TONCOUNTMENT':"吨位数分段",
                    'USEYEARSMENT':"使用年限分段",
                    'PURCHASEPRICEMENT':"新车购置价分段",
                    'LASTYEARCLAIMTIMESSEGMENT':"上年理赔次数分段",
                    'LASTYEARCLAIMSUMSSEGMENT':"上年理赔金额分段",
                    'SEATCOUNTMENTCODE':"座位数分段标识",
                    'TONCOUNTMENTCODE':"吨位数分段标识",
                    'USEYEARSMENTCODE':"使用年限分段标识",
                    'PURCHASEPRICEMENTCODE':"新车购置价分段标识",
                    'CHGOWNERFLAGNAME':"是否过户车标识",
                    'LASTYEARCLAIMTIMESSEGMENTCODE':"上年理赔次数分段标识",
                    'CAROWNER':"车主",
                    'AUTOPRODUCTRATIOMENT':"自主系数乘积分段",
                    'PAYDATE':"缴费日期",
                    'SIGNAMOUNT_BZ':"交强险签单保额",
                    'SIGNAMOUNT_A':"车辆损失险签单保额",
                    'SIGNAMOUNT_B':"三者险签单保额",
                    'SIGNAMOUNT_D':"车上人员责任险签单保额",
                    'SIGNAMOUNT_G':"全车盗抢险签单保额",
                    'SIGNAMOUNT_E':"其他险别签单保额",
                    'SIGNPREMIUM_BZ':"交强险签单保费",
                    'SIGNPREMIUM_A':"车辆损失险签单保费",
                    'SIGNPREMIUM_B':"三者险签单保费",
                    'SIGNPREMIUM_D':"车上人员责任险签单保费",
                    'SIGNPREMIUM_G':"全车盗抢险签单保费",
                    'SIGNPREMIUM_E':"其他险别签单保费",
                    'TAX_BZ':"交强险税额",
                    'TAX_A':"车辆损失险税额",
                    'TAX_B':"三者险税额",
                    'TAX_D':"车上人员责任险税额",
                    'TAX_G':"全车盗抢险税额",
                    'TAX_E':"其他险别税额",
                    'SIGNPREMIUMOFTAX_BZ':"交强险签单保费（含税）",
                    'SIGNPREMIUMOFTAX_A':"车辆损失险签单保费（含税）",
                    'SIGNPREMIUMOFTAX_B':"三者险签单保费（含税）",
                    'SIGNPREMIUMOFTAX_D':"车上人员责任险签单保费（含税）",
                    'SIGNPREMIUMOFTAX_G':"全车盗抢险签单保费（含税）",
                    'SIGNPREMIUMOFTAX_E':"其他险别签单保费（含税）",
                    'COMMCHARGE_BZ':"交强险手续费金额",
                    'COMMCHARGE_A':"车辆损失险手续费金额",
                    'COMMCHARGE_B':"三者险手续费金额",
                    'COMMCHARGE_D':"车上人员责任险手续费金额",
                    'COMMCHARGE_G':"全车盗抢险手续费金额",
                    'COMMCHARGE_E':"其他险别手续费金额",
                    'BASEPERFORMANCE_BZ':"交强险绩效金额",
                    'BASEPERFORMANCE_A':"车辆损失险绩效金额",
                    'BASEPERFORMANCE_B':"三者险绩效金额",
                    'BASEPERFORMANCE_D':"车上人员责任险绩效金额",
                    'BASEPERFORMANCE_G':"全车盗抢险绩效金额",
                    'BASEPERFORMANCE_E':"其他险别绩效金额",
                    'REGISTNUMBER_BZ':"交强险立案件数",
                    'REGISTNUMBER_A':"车辆损失险立案件数",
                    'REGISTNUMBER_B':"三者险立案件数",
                    'REGISTNUMBER_D':"车上人员责任险立案件数",
                    'REGISTNUMBER_G':"全车盗抢险立案件数",
                    'REGISTNUMBER_E':"其他险别立案件数",
                    'SUMPAY_BZ':"交强险已决赔款",
                    'SUMPAY_A':"车辆损失险已决赔款",
                    'SUMPAY_B':"三者险已决赔款",
                    'SUMPAY_D':"车上人员责任险已决赔款",
                    'SUMPAY_G':"全车盗抢险已决赔款",
                    'SUMPAY_E':"其他险别已决赔款",
                    'SUMPAYDIRECTLYFEE_BZ':"交强险已决直接理赔费用",
                    'SUMPAYDIRECTLYFEE_A':"车辆损失险已决直接理赔费用",
                    'SUMPAYDIRECTLYFEE_B':"三者险已决直接理赔费用",
                    'SUMPAYDIRECTLYFEE_D':"车上人员责任险已决直接理赔费用",
                    'SUMPAYDIRECTLYFEE_G':"全车盗抢险已决直接理赔费用",
                    'SUMPAYDIRECTLYFEE_E':"其他险别已决直接理赔费用",
                    'SUMPAYINDIRFEE_BZ':"交强险已决赔款含直费",
                    'SUMPAYINDIRFEE_A':"车辆损失险已决赔款含直费",
                    'SUMPAYINDIRFEE_B':"三者险已决赔款含直费",
                    'SUMPAYINDIRFEE_D':"车上人员责任险已决赔款含直费",
                    'SUMPAYINDIRFEE_G':"全车盗抢险已决赔款含直费",
                    'SUMPAYINDIRFEE_E':"其他险别已决赔款含直费",
                    'OUTSTANDINGPAY_BZ':"交强险未决赔款",
                    'OUTSTANDINGPAY_A':"车辆损失险未决赔款",
                    'OUTSTANDINGPAY_B':"三者险未决赔款",
                    'OUTSTANDINGPAY_D':"车上人员责任险未决赔款",
                    'OUTSTANDINGPAY_G':"全车盗抢险未决赔款",
                    'OUTSTANDINGPAY_E':"其他险别未决赔款",
                    'OUTSTANDINGDIRECTLYFEE_BZ':"交强险未决直接理赔费用",
                    'OUTSTANDINGDIRECTLYFEE_A':"车辆损失险未决直接理赔费用",
                    'OUTSTANDINGDIRECTLYFEE_B':"三者险未决直接理赔费用",
                    'OUTSTANDINGDIRECTLYFEE_D':"车上人员责任险未决直接理赔费用",
                    'OUTSTANDINGDIRECTLYFEE_G':"全车盗抢险未决直接理赔费用",
                    'OUTSTANDINGDIRECTLYFEE_E':"其他险别未决直接理赔费用",
                    'OUTSTANDINGPAYINDIRFEE_BZ':"交强险未决赔款含直费",
                    'OUTSTANDINGPAYINDIRFEE_A':"车辆损失险未决赔款含直费",
                    'OUTSTANDINGPAYINDIRFEE_B':"三者险未决赔款含直费",
                    'OUTSTANDINGPAYINDIRFEE_D':"车上人员责任险未决赔款含直费",
                    'OUTSTANDINGPAYINDIRFEE_G':"全车盗抢险未决赔款含直费",
                    'OUTSTANDINGPAYINDIRFEE_E':"其他险别未决赔款含直费",
                    'INSUREDADDRESS':"被保险人地址",
                    'AUTOPRICERATIO':"自主定价系数",
                    'BENCHMARKPREMIUM':"基准保费(含税)",
                    'BENCHMARKPREMIUMNOTAX':"基准保费(不含税)",
                    'EARNEDPREMIUMS':"满期保费",
                    'ORIPAYDATE':"原始保单缴费日期",
                    'ORIAPPDATE':"原始保单投保日期",
                    'QUALITYNUM':"核定载质量",
                    'APPLIIDENTIFYTYPE':"投保人证件类型",
                    'APPLIIDENTIFYNUMBER':"投保人证件号",
                    'APPLIADDRESS':"投保人地址",
                    'INSUREDIDENTIFYTYPE':"被保险人证件类型",
                    'INSUREDIDENTIFYNUMBER':"被保险人证件号",
                    'CARIDENTIFYTYPE':"车主证件类型",
                    'CARIDENTIFYNO':"车主证件号码",
                    'CARADDRESS':"车主地址",
                    'CARSERIESCODE':"车系代码",
                    'CARSHIPPAYTAX':"车船税金额",
                    'MAPPINGTOWNCODE':"MAPPINGTOWNCODE",
                    'NEWCARFLAG':"新旧车",
                    'IFJOINSALE':"是否联合销售",
                    'JYPOLICYNO':"驾意险保单号",
                    'JYPREMIUM':"驾意险保费",
                    'JYAMOUNT':"驾意险保额",
                    'SUMPAY_DLJY':"道路救援金额",
                    'SUMPAY_AQJC':"安全检测金额",
                    'SUMPAY_DJ':"代驾金额",
                    'SUMPAY_DWSJ':"代为送检金额",
                    'SUMPAY_ZZFW':"增值服务金额",
                    'AIRBAGCOUNT':"气囊数",
                    'SIGNAMOUNT_F':"车上人员（乘客）责任险签单保额",
                    'SIGNAMOUNT_S':"商业险签单保额",
                    'CB_NUMBER':"车上人员（乘客）责任险承保座位数",
                    'SIGNPREMIUM_S':"商业险签单保费",
                    'SIGNPREMIUM_F':"车上人员（乘客）责任险签单保费",
                    'TAX_S':"商业险税额",
                    'TAX_F':"车上人员（乘客）责任险税额",
                    'COMMCHARGE_S':"商业险手续费金额",
                    'COMMCHARGE_F':"车上人员（乘客）责任险手续费金额",
                    'SIGNPREMIUMOFTAX_S':"商业险签单保费（含税）",
                    'SIGNPREMIUMOFTAX_F':"车上人员（乘客）责任险签单保费（含税）",
                    'REGISTNUMBER_S':"商业险立案件数",
                    'REGISTNUMBER_F':"车上人员（乘客）责任险立案件数",
                    'REGISTNUMBER0':"立案件数（不含零注拒）",
                    'REGISTNUMBER_BZ0':"交强险立案件数（不含零注拒）",
                    'REGISTNUMBER_A0':"车辆损失险立案件数（不含零注拒）",
                    'REGISTNUMBER_B0':"三者险立案件数（不含零注拒）",
                    'REGISTNUMBER_S0':"商业险立案件数（不含零注拒）",
                    'REGISTNUMBER_D0':"车上司机责任险立案件数（不含零注拒）",
                    'REGISTNUMBER_F0':"车上乘客责任险立案件数（不含零注拒）",
                    'REGISTNUMBER_G0':"全车盗抢险立案件数（不含零注拒）",
                    'REGISTNUMBER_E0':"其他险别立案件数（不含零注拒）",
                    'SUMPAY_S':"商业险已决赔款",
                    'SUMPAY_F':"车上人员（乘客）责任险已决赔款",
                    'SUMPAYINDIRFEE_S':"商业险已决赔款含直费",
                    'SUMPAYINDIRFEE_F':"车上人员（乘客）责任险已决赔款含直费",
                    'SUMPAYINDIRFEE0':"已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_BZ0':"交强险已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_A0':"车辆损失险已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_B0':"三者险已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_S0':"商业险已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_D0':"车上司机责任险已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_F0':"车上乘客责任险已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_G0':"全车盗抢险已决赔款含直费（不含零注拒）",
                    'SUMPAYINDIRFEE_E0':"其他险别已决赔款含直费（不含零注拒）",
                    'OUTSTANDINGPAY_S':"商业险未决赔款",
                    'OUTSTANDINGPAY_F':"车上人员（乘客）责任险未决赔款",
                    'OUTSTANDINGPAYINDIRFEE_S':"商业险未决赔款含直费",
                    'OUTSTANDINGPAYINDIRFEE_F':"车上人员（乘客）责任险未决赔款含直费",
                    'BENCHMARKPREMIUMNOTAX_BZ':"交强险基准保费",
                    'BENCHMARKPREMIUMNOTAX_A':"车辆损失险基准保费",
                    'BENCHMARKPREMIUMNOTAX_B':"三者险基准保费",
                    'BENCHMARKPREMIUMNOTAX_S':"商业险基准保费",
                    'BENCHMARKPREMIUMNOTAX_D':"车上人员（司机）责任险基准保费",
                    'BENCHMARKPREMIUMNOTAX_F':"车上人员（乘客）责任险基准保费",
                    'BENCHMARKPREMIUMNOTAX_G':"全车盗抢险基准保费",
                    'BENCHMARKPREMIUMNOTAX_E':"其他险别基准保费",
                    'BENCHMARKPREMIUM_BZ':"交强险基准保费(含税)",
                    'BENCHMARKPREMIUM_A':"车辆损失险基准保费(含税)",
                    'BENCHMARKPREMIUM_B':"三者险基准保费(含税)",
                    'BENCHMARKPREMIUM_S':"商业险基准保费(含税)",
                    'BENCHMARKPREMIUM_D':"车上人员（司机）责任险基准保费(含税)",
                    'BENCHMARKPREMIUM_F':"车上人员（乘客）责任险基准保费(含税)",
                    'BENCHMARKPREMIUM_G':"全车盗抢险基准保费(含税)",
                    'BENCHMARKPREMIUM_E':"其他险别基准保费(含税)",
                    'EARNEDPREMIUMS_BZ':"交强险满期保费",
                    'EARNEDPREMIUMS_A':"车辆损失险满期保费",
                    'EARNEDPREMIUMS_B':"三者险满期保费",
                    'EARNEDPREMIUMS_S':"商业险满期保费",
                    'EARNEDPREMIUMS_D':"车上人员（司机）责任险满期保费",
                    'EARNEDPREMIUMS_F':"车上人员（乘客）责任险满期保费",
                    'EARNEDPREMIUMS_G':"全车盗抢险满期保费",
                    'EARNEDPREMIUMS_E':"其他险别满期保费",
                    'YIZHUAN':"已赚车年",
                    'YIZHUAN_BZ':"交强险已赚车年",
                    'YIZHUAN_A':"车辆损失险已赚车年",
                    'YIZHUAN_B':"三者险已赚车年",
                    'YIZHUAN_S':"商业险已赚车年",
                    'YIZHUAN_D':"车上人员（司机）责任险已赚车年",
                    'YIZHUAN_F':"车上人员（乘客）责任险已赚车年",
                    'YIZHUAN_G':"全车盗抢险已赚车年",
                    'YIZHUAN_E':"其他险别已赚车年",
                    'JQNCD':"交强险NCD",
                    'SYNCD':"商业险NCD",
                    'JQFDYY':"交强险浮动原因",
                    'JQBFDYY':"交强险不浮动原因",
                    'SYFDYY':"商业险浮动原因",
                    'SYBFDYY':"商业险不浮动原因",
                    'SUMPAYINDIRFEE_C0':"车上人员责任险已决赔款含直费（不含零注拒）",
                    'INSUREYEARS':"商业险连续承保年数",
                    'PRISCORE':"整车评分",
                    'PRITRANCHE':"整车分档",
                    'CLAIMTIMES':"商业险连续承保期间出险次数",
                    'TEAMTYPEID':"团队类型",
                    'STAFFFLAG':"是否员工(0否,1是)",
                    'GROUPCUSTOMERNAME':"团单名称",
                    'GROUPCUSTOMERTYPE':"团单类型",
                    'JC_POLICYNO':"家财险保单号",
                    'JC_SUMAMOUNT':"家财险保额",
                    'JC_SUMNOTAXPREMIUM':"家财险保费",
                    'JOINTSALESFLAG':"是否联合销售字段(0否,1是)"
}
# 替换列名
dfXL.rename(columns=column_mapping, inplace=True)
#------------------------------------------------------------------------------------------------------------------------------------------------------险类数据预处理

#Table5 保费进展
#table5 三年保费增速计算---------------------------------------------------------------------------

Interval=[1, 2, 3, 4,5]
filtered_df = dfXL[dfXL['月份'].isin(Interval)]
pivotTable11Proportion = filtered_df.pivot_table(
    values='签单保费',
    index=['省级业务归属区划','险种'],
    columns=['年份'],
    aggfunc='sum',
    fill_value=0,
    margins=True,
    margins_name='合计'
)

from tabulate import tabulate
province_sums = pivotTable11Proportion.groupby(level=0).sum()

# 创建一个新的 DataFrame 来存储合计行
summary_rows = pd.DataFrame(columns=pivotTable11Proportion.columns)

# 创建一个新的 DataFrame 来存储结果
combined_df = pd.DataFrame()

# 遍历每个省份，插入合计行
for province in province_sums.index:
    # 提取现有的数据
    temp_df = pivotTable11Proportion.xs(province, level=0, drop_level=False)
    # 计算合计
    province_sum = temp_df.sum(axis=0)
    # 创建合计行
    summary_row = pd.DataFrame([province_sum], index=pd.MultiIndex.from_tuples([(province, '合计')]))
    # 合并临时数据和合计行
    temp_combined_df = pd.concat([temp_df, summary_row])
    # 合并所有数据
    combined_df = pd.concat([combined_df, temp_combined_df])

combined_df['同比增速'] = (combined_df[2024] - combined_df[2023]) / combined_df[2023] * 100

# 格式化增速列显示为百分比并保留两位小数
combined_df['同比增速'] = combined_df['同比增速'].apply(lambda x: f'{x:.2f}%')
province_order = ['河南省', '内蒙古', '黑龙江省', '合计']

# 按照指定顺序排序
combined_df = combined_df.loc[province_order]
print(combined_df)
#table5 三年保费增速计算---------------------------------------------------------------------------
#table5 预计目标计算---------------------------------------------------------------------------
def process_data(path, regions, interval):
    data = []

    # 定义月份列名，选择指定区间的月份
    date_range = [f'{i}月' for i in interval]

    # 对每个地区的数据进行处理
    for region in regions:
        dfsn = pd.read_excel(path, sheet_name=f'涉农商险（{region}）', header=2, index_col=[0, 2])
        dfct = pd.read_excel(path, sheet_name=f'传统商险（{region}）', header=2, index_col=[0, 2])

        total_sn = dfsn.loc[('车险', '小计'), date_range].sum().sum()
        total_ct = dfct.loc[('车险', '小计'), date_range].sum().sum()
        total = total_sn + total_ct

        sc_sn = dfsn.loc[('车险', '商车'), date_range].sum().sum()
        sc_ct = dfct.loc[('车险', '商车'), date_range].sum().sum()
        sc_total = sc_sn + sc_ct

        jq_sn = dfsn.loc[('车险', '交强'), date_range].sum().sum()
        jq_ct = dfct.loc[('车险', '交强'), date_range].sum().sum()
        jq_total = jq_sn + jq_ct

        # data.append([region, total, sc_total, jq_total])
        data.append([region, '交强险', jq_total * 10000])
        data.append([region, '商业险', sc_total * 10000])
        data.append([region, '合计', total * 10000])

    # 创建DataFrame
    # result_df = pd.DataFrame(data, columns=['省级归属区划', '合计', '商业险', '交强险'])
    # 创建DataFrame
    result_df = pd.DataFrame(data, columns=['省级归属区划', '险种', '预计目标'])

    # 设置多级索引
    result_df.set_index(['省级归属区划', '险种'], inplace=True)
    # result_df[['合计', '商业险', '交强险']] *= 10000
    result_df = result_df.rename(index={
        '河南': '河南省',
        '黑龙江': '黑龙江省'
    })
    
    
    
    return result_df

# 示例使用
path = r"F:\\pandas\\Oracle\\firatCAR\\非农商险保费.xlsx"
regions = ['河南','内蒙古',  '黑龙江']


result_df = process_data(path, regions, Interval)
print(result_df)
#table5 预计目标计算---------------------------------------------------------------------------
#table5 拼接预计目标和三年保费 计算月度达成率---------------------------------------------------------------------------
merged_df = pd.concat([combined_df, result_df], axis=1)
merged_df['月度达成率']=merged_df[2024]/merged_df['预计目标']
# 打印结果
print(merged_df)
#table5 拼接预计目标和三年保费 计算月度达成率---------------------------------------------------------------------------