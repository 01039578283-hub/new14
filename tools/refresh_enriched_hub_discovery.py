"""Update discovery metadata for the 11 hubs without changing URL inventories."""
from pathlib import Path
from urllib.parse import unquote
from datetime import datetime, timezone
from email.utils import format_datetime
import json
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT/'tools/reports/hub-enrichment'

def main():
    report=json.loads((REPORTS/'generation.json').read_text(encoding='utf-8'))
    targets={unquote(t['url']):t for t in report['targets']}
    ns='http://www.sitemaps.org/schemas/sitemap/0.9'
    ET.register_namespace('',ns)
    tree=ET.parse(ROOT/'sitemap.xml')
    before=[n.findtext('{'+ns+'}loc') for n in tree.getroot()]
    updated=[]
    for n in tree.getroot():
        url=unquote(n.findtext('{'+ns+'}loc'))
        if url in targets:
            last=n.find('{'+ns+'}lastmod')
            if last is None: last=ET.SubElement(n,'{'+ns+'}lastmod')
            last.text=report['date']; updated.append(url)
    assert set(updated)==set(targets)
    assert before==[n.findtext('{'+ns+'}loc') for n in tree.getroot()]
    with (ROOT/'sitemap.xml').open('w',encoding='utf-8',newline='\n') as stream:
        tree.write(stream,encoding='unicode',xml_declaration=True)
    ET.register_namespace('atom','http://www.w3.org/2005/Atom')
    rss=ET.parse(ROOT/'rss.xml'); channel=rss.getroot().find('channel')
    links=[n.findtext('link') for n in channel.findall('item')]
    changed=[]
    stamp=format_datetime(datetime.now(timezone.utc))
    for item in channel.findall('item'):
        url=unquote(item.findtext('link'))
        if url not in targets: continue
        page=BeautifulSoup((ROOT/targets[url]['path']).read_text(encoding='utf-8'),'html.parser')
        item.find('title').text=page.title.get_text()
        item.find('description').text=page.select_one('meta[name=description]')['content']
        item.find('pubDate').text=stamp
        changed.append(url)
    assert set(changed)==set(targets)
    channel.find('lastBuildDate').text=stamp
    assert links==[n.findtext('link') for n in channel.findall('item')]
    with (ROOT/'rss.xml').open('w',encoding='utf-8',newline='\n') as stream:
        rss.write(stream,encoding='unicode',xml_declaration=True)
    result={'sitemapUrls':len(before),'updatedHubs':len(updated),'rssItems':len(links),'updatedRss':len(changed)}
    (REPORTS/'discovery.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
    print(json.dumps(result))

if __name__=='__main__': main()
