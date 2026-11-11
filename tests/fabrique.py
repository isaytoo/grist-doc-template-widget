# Fabrique une page de test : le vrai index.html et le vrai widget.js, un faux Grist,
# et un pilote qui écrit ses constats dans la page.
import io,sys,re
racine, source_js, sortie, pilote = sys.argv[1:5]
h=io.open(racine+'/index.html',encoding='utf-8').read()
stub="""<script>
window.__options = {};
window.grist = { ready: async function(){}, onOptions: function(){}, onRecords: function(){},
  widgetApi: { getOption: async function(k){ return window.__options[k]; }, setOption: async function(k,v){ window.__options[k]=v; (window.__ecritures=window.__ecritures||[]).push(k); } },
  docApi: { listTables: async function(){ return ['Table1']; }, fetchTable: async function(){ return { id:[1], Nom:['Durand'] }; } } };
</script>"""
h=h.replace('<script src="https://docs.getgrist.com/grist-plugin-api.js"></script>', stub)
m=re.search(r'<script src="widget\.js[^"]*"></script>', h)
version=(re.search(r'widget\.js\?v=(\d+)', h) or [None,''])[1]
h=h[:m.start()]+('<script>'+io.open(source_js,encoding='utf-8').read().replace('</script>','<\\/script>')+'</script><pre id="out"></pre><script>'+'var __VERSION_INDEX__ = "'+version+'";\n'+io.open(pilote,encoding='utf-8').read()+'</script>')+h[m.end():]
io.open(sortie,'w',encoding='utf-8').write(h)
