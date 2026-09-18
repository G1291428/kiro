# -*- coding: utf-8 -*-
"""把二进制产物包进单个 HTML 文本文件。
三条下载路径：data: URI 原生下载（不依赖 JS）→ Blob 下载 → 右键另存为。
载荷只存一份（放在 <a> 的 href 里），并带可见自检。"""
import base64, hashlib, os, sys

TPL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%(title)s</title>
<style>
 *{box-sizing:border-box;margin:0;padding:0}
 body{font:15px/1.7 "Microsoft YaHei","PingFang SC",Arial,sans-serif;color:#15222e;
      background:#eaf2fa;padding:30px}
 .card{background:#fff;max-width:660px;margin:0 auto;border:1px solid #cbdaea;padding:40px 44px}
 .eyebrow{font-size:11px;font-weight:700;letter-spacing:1.6px;color:#2e9bd6;margin-bottom:13px}
 h1{font-size:24px;color:#0b2e4f;line-height:1.35;margin-bottom:7px}
 .sub{color:#5c6875;font-size:13px;margin-bottom:24px}
 .rule{height:3px;width:52px;background:#2e9bd6;margin-bottom:24px}
 .meta{display:flex;flex-wrap:wrap;gap:26px;border-top:1px solid #cbdaea;
       border-bottom:1px solid #cbdaea;padding:15px 0;margin-bottom:24px}
 .meta div span{display:block;font-size:11px;color:#8a97a3}
 .meta div b{font-size:17px;color:#0b2e4f}
 a.btn,button.btn{display:block;width:100%%;padding:17px;background:#1c5fa8;color:#fff;
      border:0;text-align:center;text-decoration:none;cursor:pointer;
      font:700 15px/1 "Microsoft YaHei",Arial,sans-serif;letter-spacing:.5px}
 a.btn:hover,button.btn:hover{background:#0b2e4f}
 button.alt{background:#fff;color:#1c5fa8;border:1px solid #1c5fa8;margin-top:10px}
 button.alt:hover{background:#eaf2fa;color:#0b2e4f}
 .msg{margin-top:13px;padding:12px 14px;font-size:12.5px;display:none;line-height:1.6}
 .msg.good{background:#eaf2fa;color:#0b2e4f;display:block}
 .msg.bad{background:#fdf1e3;color:#8a4b00;display:block}
 .chk{margin-top:20px;border-top:1px solid #cbdaea;padding-top:16px;font-size:12px;
      color:#5c6875;line-height:1.85}
 .chk b{color:#0b2e4f}
 .ok{color:#1c7a3e;font-weight:700}
 .no{color:#c0392b;font-weight:700}
 .tip{margin-top:16px;font-size:12px;color:#8a97a3;line-height:1.85}
 code{background:#f4f8fc;padding:1px 5px;font-size:11.5px;color:#1c5fa8;word-break:break-all}
 ol{margin:6px 0 0 18px}
%(extra_css)s
</style>
</head>
<body>
<div class="card">
  <div class="eyebrow">COMMERCIAL MILK SYSTEMS / CHINA FOCUS</div>
  <h1>商用自动奶泡系统 竞品分析</h1>
  <div class="sub">%(desc)s</div>
  <div class="rule"></div>
  <div class="meta">
    <div><span>页数</span><b>11 页</b></div>
    <div><span>格式</span><b>%(fmt)s</b></div>
    <div><span>体积</span><b>%(size)s</b></div>
    <div><span>检索日</span><b>2026-09-18</b></div>
  </div>

  <a class="btn" id="dl" download="%(fname)s"
     href="data:%(mime)s;base64,%(b64)s">方式一：下载 %(fname)s</a>
  <button class="btn alt" id="dl2">方式二：如果上面没反应，点这里</button>
  <div class="msg" id="msg"></div>

  <div class="chk">
    <b>文件自检</b><br>
    载荷字符数　应为 <b>%(b64len)d</b>　实际 <b><span id="clen">…</span></b>　<span id="cres">…</span><br>
    还原后大小　应为 <b>%(rawlen)d</b> 字节　实际 <b><span id="dlen">…</span></b>　<span id="dres">…</span>
  </div>

  <div class="tip">
    两个按钮都不行的话，<b>右键点"方式一"那一行 → 选"链接另存为"</b>，手动存成
    <code>%(fname)s</code>。<br>
    还不行就用命令行还原同目录的 <code>%(b64name)s</code>：<br>
    <code>base64 -d %(b64name)s &gt; %(fname)s</code>
  </div>
</div>
%(preview)s
<script>
(function(){
  var MIME = "%(mime)s", FNAME = "%(fname)s";
  var EXPECT_B64 = %(b64len)d, EXPECT_RAW = %(rawlen)d;
  var msg = document.getElementById('msg');

  function show(text, bad){ msg.className = 'msg ' + (bad ? 'bad' : 'good'); msg.innerHTML = text; }

  // 从 href 里取回载荷，顺手清掉下载通道可能插入的换行/空格
  function b64(){
    var h = document.getElementById('dl').getAttribute('href');
    var raw = h.substring(h.indexOf('base64,') + 7);
    return raw.replace(/[^A-Za-z0-9+/=]/g, '');
  }

  function bytes(){
    var s = b64(), bin = atob(s), n = bin.length, buf = new Uint8Array(n);
    for (var i = 0; i < n; i++) buf[i] = bin.charCodeAt(i);
    return buf;
  }

  // 自检：把实际情况显示出来，避免静默失败
  var okAll = true;
  try {
    var s = b64();
    document.getElementById('clen').textContent = s.length;
    var c1 = s.length === EXPECT_B64;
    document.getElementById('cres').innerHTML = c1
      ? '<span class="ok">一致</span>' : '<span class="no">不一致，文件在传输中被改动</span>';
    okAll = okAll && c1;
    var b = bytes();
    document.getElementById('dlen').textContent = b.length;
    var c2 = b.length === EXPECT_RAW;
    document.getElementById('dres').innerHTML = c2
      ? '<span class="ok">一致</span>' : '<span class="no">不一致</span>';
    okAll = okAll && c2;
    %(preview_js)s
  } catch (e) {
    okAll = false;
    document.getElementById('cres').innerHTML = '<span class="no">解码失败：' + e.message + '</span>';
  }
  if (!okAll) {
    show('载荷校验没通过，这个 HTML 在下载过程中被改写了。请改用同目录的 '
       + '<code>%(b64name)s</code> 走命令行还原。', true);
  }

  document.getElementById('dl2').addEventListener('click', function(){
    try {
      var url = URL.createObjectURL(new Blob([bytes()], {type: MIME}));
      var a = document.createElement('a');
      a.href = url; a.download = FNAME; a.style.display = 'none';
      document.body.appendChild(a); a.click();
      setTimeout(function(){ a.remove(); URL.revokeObjectURL(url); }, 5000);
      show('已触发下载，检查浏览器的下载列表。若仍为空，请右键"方式一"那一行选"链接另存为"。');
    } catch (e) {
      show('这个浏览器拦了脚本下载（' + e.message + '）。请右键"方式一"那一行选"链接另存为"。', true);
    }
  });
})();
</script>
</body>
</html>
"""

PREVIEW_HTML = """<div class="viewer"><iframe id="pv" title="预览"></iframe></div>"""
PREVIEW_JS = """document.getElementById('pv').src =
      URL.createObjectURL(new Blob([b], {type: MIME}));"""
PREVIEW_CSS = """
 .viewer{max-width:1180px;margin:24px auto 0;background:#fff;border:1px solid #cbdaea;padding:10px}
 .viewer iframe{width:100%;height:78vh;border:0;display:block}
"""

JOBS = [
    ("milk-system-competitive-analysis.pptx", "download-pptx.html",
     "application/vnd.openxmlformats-officedocument.presentationml.presentation",
     "PowerPoint", "可编辑的 11 页演示文稿，科研蓝白风格，含 23 张品牌官方原图"),
    ("milk-system-competitive-analysis.pdf", "download-pdf.html",
     "application/pdf", "PDF", "同内容 PDF，打开本页即可直接翻阅全部 11 页"),
]

root = sys.argv[1] if len(sys.argv) > 1 else "."
for src, out, mime, fmt, desc in JOBS:
    raw = open(os.path.join(root, src), "rb").read()
    b64 = base64.b64encode(raw).decode()
    is_pdf = src.endswith(".pdf")
    html = TPL % {
        "title": "下载 " + src, "desc": desc, "fmt": fmt, "fname": src,
        "size": "%.2f MB" % (len(raw) / 1048576), "mime": mime,
        "b64name": src + ".b64.txt", "b64": b64,
        "b64len": len(b64), "rawlen": len(raw),
        "preview": PREVIEW_HTML if is_pdf else "",
        "preview_js": PREVIEW_JS if is_pdf else "",
        "extra_css": PREVIEW_CSS if is_pdf else "",
    }
    op = os.path.join(root, out)
    open(op, "w", encoding="utf-8", newline="\n").write(html)
    print("%-22s %6.2f MB   载荷 %d 字符 / 还原 %d 字节   sha256 %s"
          % (out, os.path.getsize(op) / 1048576, len(b64), len(raw),
             hashlib.sha256(raw).hexdigest()[:16]))
