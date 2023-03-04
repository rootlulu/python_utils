# excel_tuils
the excel_utils respects the workbook as the main object for using.
---

### For Example
If there is a workbook like the blew:

| A | B | C |
| ------ | ------ | ------ |
| name | sex | age |
| xiaomi | female | 20 |


Yield default is the below in to_dict():
{"name": "xiaomi", "sex": "female", "age": 20}
Yield one by one with the col_name in to_dict(show_col_names=True): 
{"A": "name", "B": "sex", "C": "age"}, {"A": "xiaomi", "B": "female", "C": 20}
Yield one by one with the col_name in to_dict(col_mapping={"name": "姓名"}): 
{"姓名": "xiaomi", "sex": "female", "age": 20}

Append with dict:
append({"name": "lulu", "sex": "male", "age": 20})
Append with list:
append(["lulu", "male",  20])
Append with whole row style:
append({"name": "lulu", "sex": "male", "age": 20, style={"color": "000FFF", "size":19})
Append with all cells style:
append({"name": "lulu", "sex": "male", "age": 20, style=[{"color": "000FFF", "size":19}, {"color": "000999", "size":99}, {"color": "0F0F0F", "size":32}])

### The excel_utils supporting for these features.
1. Operate the excel with the worrbook. Open it if it existed, create it otherwise. You can create the excel with a existed template.
2. Open the excel with a headers_idx, and then reading operation will jump over the rows whose index small than headers_idx.
3. Read excel with to_dict func and append row with append func.
4. Assign the max col when reading. The headers is the first row default and it's the headers, return the dict the other rows as the value.
5.Assign whether display the col name: A, B, C, D.
6. 
7. Use list and dict to append the row to current workbook. If it's the dict. which must be the key valu paris corresponding to headers and values.
8. Specify the style while writing.
9. Set the style in col, row or cell independently.

I know, there are a lot fo problems in my codes. Such as the to_dict would not
function in the scene what some col_names is the same. Contribute for it.

### excel_utils支持一下特性。
1. 隐藏工作簿，通过工作表操作excel。文件存在时打开，不存在时创建，创建excel时可以指定模板
2. 打开excel可以指定headers_idx，读取时从headers_idx开始读取，跳过之前行.
3. 通过to_dict读取excel和通过append追加行
4. 读取excel时可以指定读取的最大列。默认以第一行为表头，剩余行为值返回字典
5. 读取excel时可以指定是否显示col_name，即excel的表头：A，B，C，D
6. 读取excel时可以指定是否替换col值(其实为headers值)。替换col可以在返回dict时替换表头
7. 写入excel时可以通过list和dict追加行。dict写入方式为和值的对应关系
8. 写入excel时可以指定样式。单个字典为整行样式，列表字典为和值对应的样式
9. 可以单独设置行，列，单元格的样式

更多功能请阅读源码了解。


