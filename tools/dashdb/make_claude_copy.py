"""Make the Claude public-page copy of a dashboard from the GitHub Pages copy:
inline the background image and drop the robots meta tag.
usage: python3 make_claude_copy.py <repo>/<proj>/index.html <out.html>"""
import sys, os, base64
src, out = sys.argv[1], sys.argv[2]
h = open(src, encoding='utf-8').read()
img = open(os.path.join(os.path.dirname(src), 'bg.jpg'), 'rb').read()
h = h.replace('<meta name="robots" content="noindex,nofollow">', '', 1)
h = h.replace('url("bg.jpg")', 'url("data:image/jpeg;base64,' + base64.b64encode(img).decode() + '")')
open(out, 'w', encoding='utf-8').write(h)
print('wrote', out)
