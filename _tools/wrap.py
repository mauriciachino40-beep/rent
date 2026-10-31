import sys
h=open(sys.argv[1],encoding='utf-8').read()
head='<!doctype html>\n<html lang="zh-CN">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n<meta name="robots" content="noindex,nofollow">\n<meta name="format-detection" content="telephone=no">\n'
i=h.index('<style>'); j=h.index('</style>')+len('</style>')
out=head+h[:j]+'\n<style>:root{color-scheme:light dark}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n</head>\n<body>\n'+h[j:]+'\n</body>\n</html>\n'
open(sys.argv[2],'w',encoding='utf-8').write(out)
