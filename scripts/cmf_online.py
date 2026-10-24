

"""Descarga a solicitud del usuario desde dominios oficiales CMF."""
from html.parser import HTMLParser
from urllib.parse import urljoin, urlparse, unquote
from urllib.request import Request, build_opener, HTTPRedirectHandler
from datetime import date
from pathlib import PurePosixPath
import re
import calendar
import json
from scripts.importar_cmf import MESES

INDICE='https://www.cmfchile.cl/portal/estadisticas/626/w4-propertyvalue-28910.html'
PATRON='Reporte Mensual de Información Financiera del Sistema Bancario'
MAX_BYTES=15*1024*1024

def validar_url(url):
    u=urlparse(url)
    if u.scheme!='https' or u.hostname not in ('www.cmfchile.cl','cmfchile.cl') or u.username or u.password or u.port not in (None,443):
        raise ValueError('Solo se permiten enlaces HTTPS del dominio oficial de la CMF.')
    return url

class RedireccionOficial(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        validar_url(newurl)
        return super().redirect_request(req,fp,code,msg,headers,newurl)

def descargar(url):
    validar_url(url)
    req=Request(url,headers={'User-Agent':'RadiografiaBancaria/3.0 (consulta de datos publicos CMF)'})
    with build_opener(RedireccionOficial()).open(req,timeout=20) as respuesta:
        validar_url(respuesta.geturl())
        contenido=respuesta.read(MAX_BYTES+1)
    if len(contenido)>MAX_BYTES:
        raise ValueError('La respuesta supera el tamaño permitido.')
    return contenido

class Enlaces(HTMLParser):
    def __init__(self):
        super().__init__(); self.links=[]; self.actual=None
    def handle_starttag(self,tag,attrs):
        if tag=='a': self.actual=[dict(attrs).get('href',''),'']
    def handle_data(self,data):
        if self.actual is not None: self.actual[1]+=data
    def handle_endtag(self,tag):
        if tag=='a' and self.actual is not None:
            self.links.append(tuple(self.actual)); self.actual=None

def enlaces(html):
    parser=Enlaces(); parser.feed(html); return parser.links

def catalogo_desde_html(html):
    resultado={}
    for href,titulo in enlaces(html):
        titulo=' '.join(titulo.split())
        if PATRON.lower() not in titulo.lower() or not re.search(r'w4-article-\d+\.html',href):
            continue
        m=re.search(r'('+'|'.join(MESES)+r')\s+(20\d{2})',titulo.upper())
        if not m: continue
        mes,anio=MESES[m[1]],int(m[2]); fecha=f'{anio}-{mes:02d}-{calendar.monthrange(anio,mes)[1]}'
        if fecha>date.today().isoformat(): continue
        url=validar_url(urljoin(INDICE,href))
        resultado[fecha]={'fecha':fecha,'titulo':titulo,'url':url}
    return sorted(resultado.values(),key=lambda r:r['fecha'],reverse=True)

def buscar_reportes():
    registros=catalogo_desde_html(descargar(INDICE).decode('utf-8',errors='replace'))
    if not registros: raise ValueError('La página de CMF no devolvió un catálogo reconocible. Usa la carga manual.')
    return registros

def descargar_reporte(registro):
    html=descargar(registro['url']).decode('utf-8',errors='replace')
    articulo=re.search(r'w4-article-(\d+)\.html',registro['url'])
    if not articulo: raise ValueError('El enlace no es una publicación mensual reconocible.')
    # Evitar anexos de otros artículos enlazados en la misma página.
    for href,_ in enlaces(html):
        url=urljoin(registro['url'],href)
        ruta=unquote(urlparse(url).path)
        if re.search(r'articles-'+articulo[1]+r'_recurso_\d+\.xlsx$',ruta,re.I):
            contenido=descargar(url)
            if not contenido.startswith(b'PK'): raise ValueError('CMF no devolvió un archivo XLSX válido.')
            return PurePosixPath(ruta).name,contenido
    raise ValueError('No se encontró el Excel XLSX de esta publicación. Descárgalo manualmente desde la CMF.')