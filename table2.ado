* Description: output the results of the table command to Stata interface, Word and LaTeX
* Author: Meiting Wang, doctor, Institute for Economic and Social Research, Jinan University
* Email: wangmeiting92@gmail.com
* Created on Aug 16, 2020

program define table2, rclass
version 16
syntax varlist(max=2 numeric) [if/] [in] [aw fw/] [using], Contents(string) [REPLACE APPEND Format(string) ROW COLumn LISTwise EQLabels(passthru) COLLabels(passthru) VARLabels(passthru) VARWidth(passthru) MODELWidth(passthru) COMPRESS TItle(passthru) Alignment(string) PAGE(string) WIDTH(passthru)]
/*
*-实例
*--共同部分
sysuse auto.dta, clear
pr drop _all
table2 foreign, c(n) //分组计数
table2 foreign, c(n mean(price) sd(price) mean(trunk) sd(trunk)) //分组计算统计量
table2 foreign, c(n mean(price) sd(price) mean(trunk) sd(trunk)) row //额外报告行方向总体上计算的统计量
table2 foreign, c(n mean(price) sd(price) mean(trunk) sd(trunk)) row list //计算统计量时不会考虑包含缺漏值的观测值

table2 foreign rep78, c(n) //分组计数
table2 foreign rep78, c(n mean(price) sd(price) mean(trunk) sd(trunk)) //分组计算统计量
table2 foreign rep78, c(n mean(price) sd(price) mean(trunk) sd(trunk)) row //额外报告行方向总体上计算的统计量
table2 foreign rep78, c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col //额外报告列方向总体上计算的统计量
table2 foreign rep78, c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col list //计算统计量时不会考虑包含缺漏值的观测值

table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) f(0 2 2 2 2) row col //设置数值格式
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col eql(domestic foreign Total) //自定义报告的行方程名
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col varl(mean(price) price_m sd(price) price_sd mean(trunk) trunk_m sd(trunk) trunk_sd) //自定义行名
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col coll("very bad" bad general good "very good" Total) //自定义列名
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col compress //将表格压缩展示
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col varw(11) //自定义第一列的宽度(空格数)
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col compress varw(12) modelw(10) //将第二列及之后列的宽度设定为10个空格
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col compress varw(12) modelw(10 15 20 20 20 20) //为第二列及之后列分别自定义宽度
table2 foreign rep78 , c(n mean(price) sd(price) mean(trunk) sd(trunk)) row col ti(This is a title) //自定义表格标题

*--Word部分
table2 foreign rep78 using Myfile.rtf, replace c(n mean(price) sd(price) mean(trunk) sd(trunk)) f(0 2 2 2 2) row col ti(This is a title)  //将结果输出至Word

*--LaTeX部分
table2 foreign rep78 using Myfile.tex, replace c(n mean(price) sd(price) mean(trunk) sd(trunk)) f(0 2 2 2 2) row col ti(This is a title) //将结果输出至LaTeX
table2 foreign rep78 using Myfile.tex, replace c(n mean(price) sd(price) mean(trunk) sd(trunk)) f(0 2 2 2 2) row col ti(This is a title) a(math) //设置列格式为数学格式(也为默认列格式)
table2 foreign rep78 using Myfile.tex, replace c(n mean(price) sd(price) mean(trunk) sd(trunk)) f(0 2 2 2 2) row col ti(This is a title) a(dot) //设置列格式为小数点对齐
table2 foreign rep78 using Myfile.tex, replace c(n mean(price) sd(price) mean(trunk) sd(trunk)) f(0 2 2 2 2) row col ti(This is a title) page(amsmath) //引入额外的宏包(无论怎么样都会引入array和booktabs宏包)
table2 foreign rep78 using Myfile.tex, replace c(n mean(price) sd(price) mean(trunk) sd(trunk)) f(0 2 2 2 2) row col ti(This is a title) width(\textwidth) //设置表格宽度为版心宽度

*该命令结果可以用table命令进行验证
table foreign, c(freq)
table foreign, c(freq mean price sd price mean trunk sd trunk)
table foreign, c(freq mean price sd price mean trunk sd trunk) row

table foreign rep78, c(freq)
table foreign rep78, c(freq mean price sd price mean trunk sd trunk)
table foreign rep78, c(freq mean price sd price mean trunk sd trunk) row
table foreign rep78, c(freq mean price sd price mean trunk sd trunk) row col

*-编程思想或注意事项
1. 利用了自己编写的space_rm、wmtstr和mat_cagn程序
2. 通过count if获得了每个区块的观测值
3. 通过tabstat获得了每个区块的统计量
4. [if] [in] [aw fw/]参照了tabstat, 其中的权重不会影响到n
5. 每一区块的n表示count if `by1'==cat1(一个分类变量)，或count if `by1'==cat1 & `by2'==cat2(两个分类变量)
6. 每一区块的其他统计量均是在`by1'==cat1(一个分类变量)，或`by1'==cat1 & `by2'==cat2(两个分类变量)框架下进行计算的
7. 输出时采用esttab(estout)框架——输出矩阵
*/

*总体初步处理
local in_cell_num: word count `contents'
if "`alignment'" == "" {
	local alignment "math"
}

*程序报错信息
local stat_info "((\bn\b)|([A-Za-z_]\w*\(\s*[A-Za-z_]\w*\s*\)))" //匹配类似n、mean(price)、sd(mpg)、p1(rep78)等
if ~ustrregexm("`contents'","^\s*`stat_info'(\s+`stat_info')*\s*$") {
	dis "{error:wrong contents statement}"
	exit
} //对整体的contents进行限定

if "`format'" != "" {
	local format_num: word count `format'
	if ~(`format_num'==`in_cell_num') {
		dis "{error:the number of formats must be equal to the number of statistics}"
		exit
	}
}

if ("`alignment'"!="math") & ("`alignment'"!="dot") {
	dis "{error:wrong alignment setting}"
	exit
}

*处理if语句和weight语句——以兼容本程序
if "`if'" != "" {
	local if "& (`if')"
}
if "`weight'" != "" {
	local weight "[`weight'=`exp']"
}

*去掉contents括号里的空格，同时也会将括号外的长空格(连续的多个空格)压缩为单个空格
qui space_rm `contents'
local contents "`r(out_str)'"

*类别变量提取
local by1: word 1 of `varlist' //依据前面的设定，这里一定非空
local by2: word 2 of `varlist' //这里可能为空，但至多只能有两个类别变量
if ("`by2'"=="") & ("`column'"!="") {
	dis "{error:column not allowed in the case of only one categorical variable}"
	exit
}

*提取contents中所有的变量与统计量
local varlist_all = ustrregexra("`contents'","(^.*?\()|(\).*?\()|(\)[n\s]*$)|\bn\b"," ") 
if ustrwordcount("`varlist_all'") > 1 {
	qui wmtstr `varlist_all'
	local varlist_all "`r(wmtstr)'"
}

if "`casewise'" != "" {
	preserve
	local varlist_all_comma = ustrregexra("`varlist_all'","\s+","\,")
	local varlist_comma = ustrregexra("`varlist'","\s+","\,")
	qui drop if missing(`varlist_all_comma') //删除contents中包含变量中有缺漏值的观测值
	qui drop if missing(`varlist_comma') //删除by1 by2有缺漏值的观测值
}

local stat_all = ustrregexra("`contents'","\(\s*[A-Za-z_]\w*\s*\)","") 
if ustrwordcount("`stat_all'") > 1 {
	qui wmtstr `stat_all'
	local stat_all "`r(wmtstr)'"
}

if "`by2'" != "" {
	*初步处理
	qui levelsof `by1', local(by1_cat) //提取类别变量1的唯一值(从小到大)
	qui levelsof `by2', local(by2_cat) //提取类别变量2的唯一值(从小到大)
	local by1_cat "`by1_cat' Total" //by1_cat添加Total
	local by2_cat "`by2_cat' Total" //by2_cat添加Total
	local by1_cat_num: word count `by1_cat' //计算类别变量1种类数目
	local by2_cat_num: word count `by2_cat' //计算类别变量2种类数目

	local roweq_f: word 1 of `by1_cat' //类别1变量的第一分类的值
	local roweq_l: word `=`by1_cat_num'-1' of `by1_cat' //类别1变量的最后分类的值
	local contents_f: word 1 of `contents' //contents的第一个“单词”
	local contents_l: word `in_cell_num' of `contents' //contents的最后一个“单词”
	local coleq_f: word 1 of `by2_cat' //类别2变量的第一分类的值
	local coleq_l: word `=`by2_cat_num'-1' of `by2_cat' //类别2变量的最后分类的值

	*构造一个`freq'矩阵，用于记录每个区块的观测值(行列名为最终输出矩阵的方程名)
	tempname freq
	mat `freq' = J(`by1_cat_num',`by2_cat_num',.)
	mat rown `freq' = `by1_cat'
	mat coln `freq' = `by2_cat'
	foreach i of local by1_cat {
		foreach j of local by2_cat {
			if "`i'"=="Total" {
				if "`j'"=="Total" {
					qui count if (`by1'!=.) & (`by2'!=.) `if' `in'
				}
				else {
					qui count if (`by1'!=.) & (`by2'==`j') `if' `in'
				}
			}
			else {
				if "`j'"=="Total" {
					qui count if (`by1'==`i') & (`by2'!=.) `if' `in'
				}
				else {
					qui count if (`by1'==`i') & (`by2'==`j') `if' `in'
				}
			}
			qui mat_cagn `freq'["`i'","`j'"] = r(N)
		}
	}

	*初步构造输出矩阵（设置其行列维度及其行列名）
	local row_eq "" //矩阵行方程名
	local row_name "" //矩阵行名
	foreach i of local by1_cat {
		local row_eq "`row_eq'`="`i' "*`in_cell_num''"
		local row_name "`row_name'`contents' "
	}

	tempname output
	mat `output' = J(`by1_cat_num'*`in_cell_num',`by2_cat_num',.)
	mat roweq `output' = `row_eq'
	mat rown `output' = `row_name'
	mat coln `output' = `by2_cat' //矩阵列名

	*对输出矩阵里的内容进行填充(包括行列Total——即完全体)
	foreach i of local by1_cat {
		foreach j of local by2_cat {
			if "`i'" == "Total" {
				if "`j'" == "Total" {
					cap tabstat `varlist_all' if (`by1'!=.) & (`by2'!=.) `if' `in' `weight', s(`stat_all') save
				}
				else {
					cap tabstat `varlist_all' if (`by1'!=.) & (`by2'==`j') `if' `in' `weight', s(`stat_all') save
				}
			}
			else {
				if "`j'" == "Total" {
					cap tabstat `varlist_all' if (`by1'==`i') & (`by2'!=.) `if' `in' `weight', s(`stat_all') save
				}
				else {
					cap tabstat `varlist_all' if (`by1'==`i') & (`by2'==`j') `if' `in' `weight', s(`stat_all') save
				}
			}
			
			if _rc == 0 {
				foreach k of local contents {
					if "`k'" == "n" {
						qui mat_cagn `output'["`i':`k'","`j'"] = `freq'["`i'","`j'"]
					}
					else {
						local stat_c = ustrregexra("`k'","\(.+\)","")
						local var_c = ustrregexra("`k'","^\w+\(|\)$","")
						qui mat_cagn `output'["`i':`k'","`j'"] = r(StatTotal)["`stat_c'","`var_c'"]
					}
				}
			}
			else if (_rc==2000) | (_rc==100) {
				foreach k of local contents {
					if "`k'" == "n" {
						qui mat_cagn `output'["`i':`k'","`j'"] = `freq'["`i'","`j'"]
					}
				}
			}
			else if _rc == 111 {
				dis "{error:some unknown variables}"
				exit 111
			}
			else if _rc == 198 {
				dis "{error:some unknown statistics}"
				exit 198
			}
			else {
				dis "{error:syntax error}"
				exit _rc
			}
		}
	}

	*依据row和column选项，从完全体矩阵中进行抽取子矩阵
	if ("`row'"=="") & ("`column'"=="") {
		mat `output' = `output'["`roweq_f':`contents_f'".."`roweq_l':`contents_l'","`coleq_f'".."`coleq_l'"]
	}
	else if ("`row'"!="") & ("`column'"=="") {
		mat `output' = `output'["`roweq_f':`contents_f'".."Total:`contents_l'","`coleq_f'".."`coleq_l'"]
	}
	else if ("`row'"=="") & ("`column'"!="") {
		mat `output' = `output'["`roweq_f':`contents_f'".."`roweq_l':`contents_l'","`coleq_f'".."Total"]
	}
	
	*format语句处理
	if "`format'" != "" {
		local format = "`format' "*`=rowsof(`output')/`in_cell_num''
		local format `""`format'""'
	}
}
else { //只有一个分类变量的情况
	*初步处理
	qui levelsof `by1', local(by1_cat) //提取类别变量1的唯一值(从小到大)
	local by1_cat "`by1_cat' Total" //by1_cat添加Total
	local by1_cat_num: word count `by1_cat' //计算类别变量1种类数目(也是最后输出矩阵的行数)

	local rown_f: word 1 of `by1_cat' //类别1变量的第一分类的值
	local rown_l: word `=`by1_cat_num'-1' of `by1_cat' //类别1变量的最后分类的值
	local contents_f: word 1 of `contents' //contents的第一个“单词”
	local contents_l: word `in_cell_num' of `contents' //contents的最后一个“单词”

	*构造一个`freq'矩阵，用于记录每个区块的观测值(在这里A=1为一个区块——cell)
	tempname freq
	mat `freq' = J(`by1_cat_num',1,.)
	mat rown `freq' = `by1_cat'
	mat coln `freq' = n
	foreach i of local by1_cat {
		if "`i'"=="Total" {
			qui count if `by1' != .
		}
		else {
			qui count if `by1' == `i'
		}
		qui mat_cagn `freq'["`i'","n"] = r(N)
	}	

	*初步构造输出矩阵（设置其行列维度及其行列名）
	tempname output
	mat `output' = J(`by1_cat_num',`in_cell_num',.)
	mat rown `output' = `by1_cat'
	mat coln `output' = `contents'

	*对输出矩阵里的内容进行填充(包括行Total——即完全体)	
	foreach i of local by1_cat {
		if "`i'" == "Total" {
			cap tabstat `varlist_all' if (`by1'!=.) `if' `in' `weight', s(`stat_all') save
		}
		else {
			cap tabstat `varlist_all' if (`by1'==`i') `if' `in' `weight', s(`stat_all') save
		}
		
		if _rc == 0 {
			foreach k of local contents {
				if "`k'" == "n" {
					qui mat_cagn `output'["`i'","`k'"] = `freq'["`i'","`k'"]
				}
				else {
					local stat_c = ustrregexra("`k'","\(.+\)","")
					local var_c = ustrregexra("`k'","^\w+\(|\)$","")
					qui mat_cagn `output'["`i'","`k'"] = r(StatTotal)["`stat_c'","`var_c'"]
				}
			}
		}
		else if (_rc==2000) | (_rc==100) {
			foreach k of local contents {
				if "`k'" == "n" {
					qui mat_cagn `output'["`i'","`k'"] = `freq'["`i'","`k'"]
				}
			}
		}
		else if _rc == 111 {
			dis "{error:some unknown variables}"
			exit 111
		}
		else if _rc == 198 {
			dis "{error:some unknown statistics}"
			exit 198
		}
		else {
			dis "{error:syntax error}"
			exit 9999
		}
	}
	
	*依据row选项，从完全体矩阵中进行抽取子矩阵
	if "`row'"=="" {
		mat `output' = `output'["`rown_f'".."`rown_l'","`contents_f'".."`contents_l'"]
	}
}


*输出结果
esttab matrix(`output',fmt(`format')), ml(,none) `eqlabels' `collabels' `varlabels' `varwidth' `modelwidth' `compress' `title'
if ustrregexm(`"`using'"',"\.rtf") {
	esttab matrix(`output',fmt(`format')) `using', ml(,none) `replace' `append' `eqlabels' `collabels' `varlabels' `varwidth' `modelwidth' `compress' `title'
}
else if ustrregexm(`"`using'"',"\.tex") {
	*LaTeX collabels专属设置
	if `"`collabels'"' == "" {
		local output_coln: coln `output'
		tokenize `"`output_coln'"'
		local i = 1
		while "``i''" != "" {
			local collabels `"`collabels'\multicolumn{1}{c}{``i''} "'
			local `i' "" //置空`i'
			local i = `i' + 1
		}
		local collabels `"collabels(`collabels')"'
	}
	else {
		local collabels = ustrregexra(`"`collabels'"',"collabels\(\s*|\s*\)","")
		tokenize `"`collabels'"'
		local i = 1
		local collabels ""
		while "``i''" != "" {
			local collabels `"`collabels'"\multicolumn{1}{c}{``i''}" "'
			local `i' "" //置空`i'
			local i = `i' + 1
		}
		local collabels `"collabels(`collabels')"'
	}
	
	*LaTeX alignment page设置
	if "`alignment'" == "math" {
		local alignment "*{`=colsof(`output')'}{>{$}c<{$}}"
		if "`page'" == "" {
			local page "page(array)"
		}
		else {
			local page "page(`page',array)"
		}
	}
	else if "`alignment'" == "dot" {
		local alignment "*{`=colsof(`output')'}{D{.}{.}{-1}}"
		if "`page'" == "" {
			local page "page(array,dcolumn)"
		}
		else {
			local page "page(`page',array,dcolumn)"
		}
	}
	local alignment "alignment(`alignment')"
	
	esttab matrix(`output',fmt(`format')) `using', ml(,none) `replace' `append' `eqlabels' `collabels' `varlabels' `varwidth' `modelwidth' `compress' `title' `alignment' `page' booktabs `width'
}


*返回值
return matrix output = `output', copy


*恢复数据集
if "`casewise'" != "" {
	restore //恢复数据集
}
end



