#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
数据统计平台 —— 一键构建脚本

背景：index.html / mobile.html 是「构建产物」（自包含单文件），
      app.js / style.css 是「可编辑源」。改了源必须重新内联，否则改动不生效。
      历史上发生过「用不完整的源 style.css 覆盖完整的内联 CSS」导致整页样式丢失的事故，
      因此本脚本内置了 CSS 覆盖率保护。

用法：
    python build.py             # 内联 app.js + style.css 进两个 HTML，并校验
    python build.py --js        # 只内联 app.js
    python build.py --css       # 只内联 style.css
    python build.py --check     # 只校验产物与源是否一致（不写入）
    python build.py --test      # 构建后额外跑 integration_test.py
    python build.py --force     # 跳过 CSS 覆盖率保护（仅在确认源 CSS 完整时使用）
"""
import io, os, re, sys, shutil, subprocess

BASE = os.path.dirname(os.path.abspath(__file__))
TARGETS = ['index.html', 'mobile.html']
JS_MARK = 'var NOTE_TEXT'          # app.js 内联块的唯一定位标记
NODE = r'C:\Users\admin\.workbuddy\binaries\node\versions\22.22.2-3\node.exe'


def rd(p):
    return io.open(os.path.join(BASE, p), encoding='utf-8').read()


def wr(p, s):
    io.open(os.path.join(BASE, p), 'w', encoding='utf-8', newline='').write(s)


def classes_of(css):
    """提取 CSS 中定义过的类名集合（含 .a.b / .a:hover 等）"""
    return set(re.findall(r'\.([a-zA-Z][\w-]*)', css))


def style_span(html):
    i = html.index('<style>') + len('<style>')
    j = html.index('</style>', i)
    return i, j


def script_span(html):
    i = html.index(JS_MARK)
    s = html.index('>', html.rindex('<script', 0, i)) + 1
    e = html.index('</script>', i)
    return s, e


def verify(html, js_src=None, css_src=None):
    """结构校验，返回 (ok, [问题...])"""
    bad = []
    if html.count('<script') != html.count('</script>'):
        bad.append('script 标签不配对 (%d/%d)' % (html.count('<script'), html.count('</script>')))
    if html.count('<style>') != 1:
        bad.append('style 块数异常 (%d)' % html.count('<style>'))
    if html.count(JS_MARK) != 1:
        bad.append('JS 定位标记出现 %d 次（应为 1）' % html.count(JS_MARK))
    if 'src="app.js"' in html or 'href="style.css"' in html:
        bad.append('仍存在外部引用（未自包含）')
    if not html.rstrip().endswith('</html>'):
        bad.append('结尾不是 </html>')
    if js_src is not None:
        s, e = script_span(html)
        if html[s:e].strip() != js_src.strip():
            bad.append('内联 JS 与 app.js 不一致')
    if css_src is not None:
        i, j = style_span(html)
        if html[i:j].strip() != css_src.strip():
            bad.append('内联 CSS 与 style.css 不一致')
    return (not bad), bad


def main():
    args = set(sys.argv[1:])
    do_js = '--js' in args or ('--css' not in args)
    do_css = '--css' in args or ('--js' not in args)
    check_only = '--check' in args
    force = '--force' in args

    js_src = rd('app.js')
    css_src = rd('style.css')

    # 语法检查（源）
    if os.path.exists(NODE):
        r = subprocess.run([NODE, '--check', os.path.join(BASE, 'app.js')],
                           capture_output=True, text=True)
        if r.returncode != 0:
            print('[中止] app.js 语法错误：\n' + (r.stderr or '')[:500])
            return 2
        print('[OK] app.js 语法检查通过')
    else:
        print('[跳过] 未找到 node，跳过语法检查')

    if do_css:
        # ---- CSS 覆盖率保护：防止源 CSS 是子集时把内联样式抹掉 ----
        # --check 模式下只报告不中止；实际写入时命中缺失即中止
        for f in TARGETS:
            html = rd(f)
            i, j = style_span(html)
            old, new = classes_of(html[i:j]), classes_of(css_src)
            missing = sorted(old - new)
            if missing:
                print('[警告] %s：源 style.css 相比现有内联样式缺少 %d 个类的定义' % (f, len(missing)))
                print('       ' + ', '.join(missing[:40]) + (' …' if len(missing) > 40 else ''))
                if not force:
                    if check_only:
                        print('       （--check 模式，仅报告）')
                    else:
                        print('[中止] 内联 CSS 会导致样式丢失。确认源 CSS 完整后加 --force，或先补全 style.css。')
                        return 3
                else:
                    print('       --force 已指定，继续。')
            else:
                print('[OK] %s：源 CSS 覆盖全部 %d 个类' % (f, len(old)))

    for f in TARGETS:
        html = rd(f)
        if not check_only:
            shutil.copyfile(os.path.join(BASE, f), os.path.join(BASE, f + '.bak'))
            try:
                if do_js:
                    s, e = script_span(html)
                    html = html[:s] + '\n' + js_src + '\n' + html[e:]
                if do_css:
                    i, j = style_span(html)
                    html = html[:i] + '\n' + css_src + '\n' + html[j:]
                wr(f, html)
            finally:
                html = rd(f)   # 重新读取真实落盘内容做校验
        ok, bad = verify(html, js_src if do_js or True else None, css_src if do_css or True else None)
        size = os.path.getsize(os.path.join(BASE, f))
        tag = 'PASS' if ok else 'FAIL'
        print('[%s] %-12s %d 字节  script=%d/%d  style=%d' % (
            tag, f, size, html.count('<script'), html.count('</script>'), html.count('<style>'))) 
        for b in bad:
            print('       - ' + b)
        if not ok:
            return 1
        bak = os.path.join(BASE, f + '.bak')
        if os.path.exists(bak):
            os.remove(bak)

    if '--test' in args:
        r = subprocess.run([sys.executable, os.path.join(BASE, 'integration_test.py')],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
        tail = [l for l in (r.stdout or '').splitlines() if 'PASS' in l or 'FAIL' in l][-4:]
        print('[测试] ' + ' | '.join(tail) if tail else '[测试] 无输出')
        if r.returncode != 0:
            return 1

    print('\n完成：' + ('校验通过（未修改文件）' if check_only else '已内联并校验通过'))
    return 0


if __name__ == '__main__':
    sys.exit(main())
