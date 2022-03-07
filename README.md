# doc文档合并

## 需求

校友帮实习日记导出后根据每一天新建了一个word文档，为了方便打印，需要手动拼接成一个word文档，50篇日志就需要50次无意义的重复劳动，因此写了此脚本来快速合并文档

## 参考

- [python合并word](https://www.jianshu.com/p/a9df5e74568a)
- [利用 Python 批量合并 Word 文档](https://peppernotes.top/2020/05/pythoncombineword/)
- [实例15：用Python批量转换doc文件为docx文件](https://zhuanlan.zhihu.com/p/64189783)
- [python根据字符串中的数字进行排序](https://blog.csdn.net/fengyu7789/article/details/121766897)

## 可能存在的问题

- 每个人的文档命名方式有所不同，因此需要根据实际情况来对文档进行排序，我使用的是正则表达式找出路径中的序号
- macOS系统下可能无法使用
- 转存为`.docx`需要word应用支持

## 依赖

```
pip install pypiwin32
pip install python-docx
pip install docxcompose
```

## 改进

- 该项目建立只是为了临时解决上述需求，因为近期报告比较多，就没有对代码进行封装优化，之后有空再完善
- 代码注释也没有写得很详细

## 有用的话记得点星星🌟