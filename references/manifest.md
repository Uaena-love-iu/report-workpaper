# 清单格式 v1

使用UTF-8 JSON。图片相对于manifest目录，便于项目移动；也接受用户授权的本地绝对路径。默认不下载远程图片。

```json
{
  "version": 1,
  "as_of": "2026-11-01",
  "sheets": [{
    "label": "3.4",
    "rows": [{
      "id": "3.4-P1-S1",
      "text": "保留报告中的原句。",
      "status": "correction",
      "note": "指出具体金额、主体、期间或因果错误。",
      "suggestion": "可以替换的表述。",
      "sources": [{
        "title": "公司正式披露名称",
        "url": "https://example.com/report.pdf",
        "source_class": "company_filing",
        "kind": "pdf_crop",
        "pdf_page": 43,
        "printed_page": 42,
        "published_at": "2026-05-16",
        "collected_at": "2026-11-01",
        "image": "sources/proof.png",
        "verified": true
      }]
    }]
  }]
}
```

`verified:true`表示实际已复核，不由生成器自动赋予。URL与图片存在不代表来源真实或结论成立。

状态：`support`需要已核原截图并覆盖每句；`partial`须写缺口；`correction`须写问题和建议；`analysis`须写事实基础/推断边界；`unsupported`须写`search_log`；`deferred_ifind`须有明确`user_deferred:true`及备注。

来源类别：`company_filing`、`official`、`government`、`industry_body`、`institutional_media`、`user_database`、`archive`。不能把券商研报或个人文章标成机构媒体。

图片类型：`pdf_crop`、`web_screenshot`、`database_screenshot`、`archive_excerpt`。网页/数据库提供`location`；PDF物理页为正整数，印刷页可为字符串或明确`null`（确实没有页码）。存档摘录不能独自支撑`support`。

若一条带有多句，提供`sentence_ids`，每个来源提供`claim_ids`，确保每句单独对应。`support`要求逐句覆盖，不能把整段挂一个宽泛链接。

## 原生表格/产品表

在sheet中加入`tables`。每列有`title`、`span`（占普通列数）及可选`format`（Excel不变区域格式）。`kind:product`使用楷体；`role:series`为浅蓝系列列。表格行长度必须与列数相同。

单元格可以是数值、文本、空值，或`{"image":"products/a.png"}`。显式`{"formula":"=C2/B2"}`才作为公式；普通以等号开头的文字不执行。日期可使用`{"date":"2026-11-01"}`。金额和比例用数值，币种/比例通过格式或列名说明，缺失留空而非0。

产品名、货号、功能、原价、到手价、折扣、销量、采集日期等按实际需要列示。不同币种不直接比较折扣；同一时点数据才可计算。未核实的货号、销量及历史价格不能补造。

## 演示与兼容

`demo:true`时才接受`synthetic_demo`，并在工作表显示明显演示标记。演示数据不经过真实证据核验。

旧版`left_img`清单由`legacy_build_workpaper.py`处理。新版要求`text`，不会静默把正文转成图片。
