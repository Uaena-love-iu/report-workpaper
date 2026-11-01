# 仅更新指定工作表

1. 记录最新原文件SHA256、目标表关系和全部ZIP部件hash。确认不是锁文件、旧导出或未保存现场版本。
2. 在隔离副本中编辑目标表，原文与建议分别保存。不要删除其他表仍使用的共享媒体。
3. `patch_workbook.py`合回单独编辑的表：旧共享字符串不改；新增字符串用inlineStr；样式只追加并映射；媒体按hash去重；更新必要关系与内容类型。
4. 写回前核对表名和原文件hash。不复制新文件的全局workbook.xml、主题、外链或自定义元数据覆盖原文件。
5. 非目标表做部件级对比，旧样式做前缀不变检查。允许变化仅限目标表XML、独占drawing/关系、新图片及必要样式/内容类型登记。
6. 目标含不支持的共享drawing、图表、VML、批注、嵌入对象、表定义或条件格式依赖时，停止自动移植，改用可保留特性的宿主原生工具。不能删除特性让脚本跑通。

```text
python scripts/patch_workbook.py original.xlsx edited-single-sheet.xlsx --sheet 3.4 --out candidate.xlsx --expected-sha256 ORIGINAL_SHA --report scope.json
python scripts/verify_workbook.py candidate.xlsx --baseline original.xlsx --allowed-sheet 3.4
```

先验候选，再按授权写回。文件锁阻止写入时，不强关用户Excel或覆盖期间的新编辑。

SDK报错需比基线；旧验证器对新Office扩展的原有提示不等于新增损坏。打开候选副本抽查后不保存关闭，避免全本重存掩盖“其他表未改”的边界。
