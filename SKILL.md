---
name: zhoupu-order-import
description: 舟谱系统订单导入模板生成（自提订单+调拨订单）+ CDP自动化导入。当下单表数据需要转化为舟谱系统可导入的Excel格式，并自动导入到舟谱系统时使用。
---

# 舟谱系统订单导入（模板生成 + 自动导入）

## 完整业务流程（低温奶行业 - 短保产品）

1. **报单**：业务员/导购/分销商用企业微信智能表格提前报单（报单日→到货日）
2. **汇总**：张俊峰导出Excel，给到我填入下单表
3. **审核**：我填好后发给**郝洋**（销售经理）审核修改
4. **生成模板**：郝洋确认后，我生成舟谱导入模板

## 注意事项
- 短保产品需提前报单（报单日和到货日不同）
- 审核环节必须经过郝洋，不能跳过

---

## 概述

将客户下单表 + 价格表，按客户类型匹配对应价格，生成舟谱系统可导入的：
1. **自提订单**导入模板
2. **调拨订单**导入模板

---

## 脚本

### 1. 自提订单

```bash
python scripts/generate_order_import.py \
  --order <下单表.xlsx> \
  --price <价格表.xlsx> \
  --arrival <到货日期YYYYMMDD>
```

参数：
| 参数 | 必填 | 说明 |
|------|------|------|
| `--order` | ✅ | 下单表Excel路径 |
| `--price` | ✅ | 价格表Excel路径 |
| `--arrival` | ✅ | 到货日期YYYYMMDD |
| `--sheet` | ❌ | 工作表名关键词 |
| `--start-seq` | ❌ | 起始序号，默认21 |
| `--output` | ❌ | 输出路径 |
| `--extra-prices` | ❌ | 补充价格JSON |
| `--config` | ❌ | 客户配置JSON |

### 2. 调拨订单

```bash
python scripts/generate_transfer_order.py \
  --order <下单表.xlsx> \
  --price <价格表.xlsx> \
  --arrival <到货日期YYYYMMDD>
```

参数：
| 参数 | 必填 | 说明 |
|------|------|------|
| `--order` | ✅ | 下单表Excel路径 |
| `--price` | ✅ | 价格表Excel路径 |
| `--arrival` | ✅ | 到货日期YYYYMMDD |
| `--sheet` | ❌ | 工作表名关键词 |
| `--start-seq` | ❌ | 起始序号，默认1 |
| `--output` | ❌ | 输出路径 |

---

## 业务规则

### 自提订单
- 分销商用分销价格，门店用对应价格（永辉/沃尔玛/美联）
- 客户名称映射：易胜琳→易胜玲，谢总→宜城谢总，吾悦→永辉吾悦店等
- **源单据号格式：ZT+日期+序号（如ZT2026040801）** ← 舟谱系统要求必须有ZT前缀
- 固定值：业务员=张俊峰，部门=湖北福宝商贸有限公司，仓库=总仓

### 调拨订单
- 调拨业务员：程欢欢、刘善涛、毛辉、周运潘、田顺达、王琴、刘正宝
- 调出仓=总仓，调入仓=业务员名+仓（如程欢欢仓）
- 源单据号格式：DB+日期+序号（如DB2026040801）
- **商品条码是必填字段**，条码为空时尝试用简称关键词匹配

---

## 自动核对

生成后自动核对每个客户/业务员的订单数量合计，与下单表对比（排除合计行）。如有差异会提示。

---

## 常见问题排查

### 导入失败：成功0条，失败N条
**最常见原因：源单据号格式错误**

舟谱系统要求源单据号必须有前缀：
- 自提订单：`ZT` + 日期 + 序号（如 `ZT2026040801`）
- 调拨订单：`DB` + 日期 + 序号（如 `DB2026040801`）

排查方法：下载一个已成功导入的订单（如2026年4月3日样表），对照源单据号格式，修正脚本中的格式字符串。

### 价格缺失
某些条码在价格表中没有对应价格（尤其是美联价格）。解决方案：
1. 生成时加 `--extra-prices '{"条码":{"价格类型":数值}}'` 补充
2. 或检查下单表是否有新的SKU未录入价格表

### 商品条码浮点数问题
Excel中长条码可能被存为科学计数法，脚本已自动用 `int()` 转整数字符串。

### 下单表合计行
下单表底部常有合计行（条码为空），脚本已自动跳过。如核对数量有出入，检查是否漏跳。

### 起始序号冲突
舟谱系统已有01~20的单号，默认从21开始自提订单。如再次导入同一批客户，需用更大的序号（如41）避开冲突。

---

---

## CDP自动化导入（Mac + QQ浏览器）

模板生成后，可通过QQ浏览器CDP自动化将文件直接导入舟谱系统，无需手动操作。

### 前置条件

1. **QQ浏览器**已安装且已登录舟谱系统（portal.zhoupudata.com）
2. **Python3** + `websocket-client` 包
3. QQ浏览器需带调试端口启动

### 启动QQ浏览器（调试模式）

```bash
# 先关闭QQ浏览器
osascript -e 'quit app "QQBrowser"'
sleep 2
# 带调试端口重启
open -a QQBrowser --args --remote-debugging-port=9222
```

### 舟谱页面结构

舟谱是SPA + iframe结构：
- **主页面**：`portal.zhoupudata.com/saas/main`（外壳框架）
- **业务iframe**：`statics.zhoupudata.com/saas/erpweb/index.html`（侧边栏+内容区）

CDP操作必须在iframe的执行上下文中进行，需要用 `Page.createIsolatedWorld(frameId)` 获取。

### 自动化导入流程

```
1. CDP连接 → 获取WebSocket URL
2. 获取iframe的frameId → 创建IsolatedWorld执行上下文
3. Hover「采销管理」展开子菜单 → CDP鼠标点击目标子菜单（自提订单/调拨订单）
4. 点击「批量操作」→ Ant Design下拉菜单 → 点击「批量导入」
5. DOM.setFileInputFiles 上传xlsx文件
6. 点击「立即导入」
7. 读取结果（.ant-modal-confirm-body）
8. 关闭成功提示框 → 重复3-7导入下一个类型
```

### CDP关键操作代码（Python + websocket-client）

```python
import json, requests, websocket, time

CDP_PORT = 9222
msg_id = 0

def get_ws_url():
    tabs = requests.get(f"http://127.0.0.1:{CDP_PORT}/json/list").json()
    for t in tabs:
        if "zhoupu" in t.get("url", "").lower():
            return t["webSocketDebuggerUrl"]
    return None

def cdp(ws, method, params=None):
    global msg_id
    msg_id += 1
    m = {"id": msg_id, "method": method}
    if params: m["params"] = params
    ws.send(json.dumps(m))
    while True:
        resp = json.loads(ws.recv())
        if resp.get("id") == msg_id: return resp

# 连接
ws_url = get_ws_url()
ws = websocket.create_connection(ws_url)
cdp(ws, "DOM.enable")
cdp(ws, "Page.enable")

# 获取iframe上下文
child_frame_id = "<从Page.getFrameTree获取>"
result = cdp(ws, "Page.createIsolatedWorld", {
    "frameId": child_frame_id, "worldName": "zhoupu"
})
ctx_id = result["result"]["executionContextId"]

# 获取元素坐标并CDP鼠标点击
def click_element(ws, ctx_id, selector_js):
    result = cdp(ws, "Runtime.evaluate", {
        "expression": selector_js,
        "contextId": ctx_id, "returnByValue": True
    })
    pos = result["result"]["result"]["value"]
    if pos:
        x, y = pos["x"], pos["y"]
        cdp(ws, "Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y})
        time.sleep(0.3)
        cdp(ws, "Input.dispatchMouseEvent", {"type": "mousePressed", "x": x, "y": y, "button": "left", "clickCount": 1})
        cdp(ws, "Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x, "y": y, "button": "left", "clickCount": 1})
    return pos

# Hover采销管理 → 点击子菜单
click_element(ws, ctx_id, """
    (function() {
        var items = document.querySelectorAll('.menu-name___aaF17');
        for (var i = 0; i < items.length; i++) {
            if (items[i].textContent.trim() === '采销管理') {
                var r = items[i].getBoundingClientRect();
                return {x: r.x + r.width/2, y: r.y + r.height/2};
            }
        }
        return null;
    })()
""")
time.sleep(2)  # 等待子菜单展开

# 点击调拨订单（或自提订单）
click_element(ws, ctx_id, """
    (function() {
        var items = document.querySelectorAll('.submenu-item-title___0SKPG');
        for (var i = 0; i < items.length; i++) {
            if (items[i].textContent.trim() === '调拨订单') {
                var r = items[i].getBoundingClientRect();
                return {x: r.x + r.width/2, y: r.y + r.height/2, visible: r.width > 0};
            }
        }
        return null;
    })()
""")
time.sleep(3)  # 等待页面跳转

# 上传文件
result = cdp(ws, "DOM.performSearch", {"query": 'input[type="file"]'})
search_id = result["result"]["searchId"]
count = result["result"]["resultCount"]
if count > 0:
    nodes = cdp(ws, "DOM.getSearchResults", {"searchId": search_id, "fromIndex": 0, "toIndex": count})
    cdp(ws, "DOM.setFileInputFiles", {
        "files": ["/path/to/模板文件.xlsx"],
        "nodeId": nodes["result"]["nodeIds"][0]
    })
```

### 重要注意事项

1. **iframe frameId会变化**：每次页面刷新后需重新获取（`Page.getFrameTree`），不要硬编码
2. **CSS类名带哈希后缀**：如 `.menu-name___aaF17`，舟谱更新后可能变化，需实时检查
3. **Ant Design下拉菜单**：必须用CDP原生鼠标事件（`Input.dispatchMouseEvent`），不能用JS `.click()` 或 `element.click()`
4. **文件上传**：`DOM.setFileInputFiles` 有效，但React受控组件会显示 `files=0`，实际已成功
5. **导入间隔**：两种订单连续导入时，先关闭上一个成功提示框（按Escape），再操作下一个
6. **QQ浏览器优先**：Chrome ≥ 147 的 System Profile 限制会阻止CDP连接默认profile，QQ浏览器无此限制

### 导入路径（手动参考）

| 订单类型 | 路径 |
|----------|------|
| 自提订单 | 采销管理 → 自提订单 → 批量操作 → 批量导入 |
| 调拨订单 | 采销管理 → 调拨订单 → 批量操作 → 批量导入 |
| 车销订单 | 采销管理 → 车销订单 → 批量操作 → 批量导入 |
| 访销订单 | 采销管理 → 访销订单 → 批量操作 → 批量导入 |

---

## 上传目录（Windows手动导入用）

生成后复制到：`C:\tmp\openclaw\uploads\`

**注意**：Windows路径，直接用本地路径（如 `C:\tmp\...`），**不要用UNC网络路径**（`\\tmp\...`）。

```powershell
copy "输出文件.xlsx" "C:\tmp\openclaw\uploads\"
```

---

## 客户配置（内置）

```json
{
  "分销商": ["唐成", "黄家伟", "易胜琳", "易胜玲", "胡奎奎", "朱青峰", "谢总"],
  "永辉门店": ["吾悦", "东津", "民发"],
  "沃尔玛": ["沃尔玛"],
  "美联门店": ["檀溪美联"],
  "业务员": "张俊峰",
  "部门": "湖北福宝商贸有限公司",
  "仓库": "总仓"
}
```

价格类型映射：
- 分销商 → `分销价格`
- 永辉吾悦/东津/民发 → `永辉价格`
- 沃尔玛 → `沃尔玛价格`
- 美联檀溪 → `美联价格`
