
"""Persistencia local. El panel lee SQLite, no el CSV inicial."""
from pathlib import Path
from io import BytesIO
import hashlib
import sqlite3
import zipfile
from contextlib import contextmanager
import pandas as pd
from scripts.importar_cmf import MAPA, extraer

COLUMNAS = ['fecha', 'banco', 'tipo_entidad', 'alcance', 'archivo_fuente', *MAPA]
TRAZAS = ['fecha', 'banco', 'indicador', 'valor', 'unidad', 'archivo', 'hoja', 'celda']
NUEVAS = ['roe_pct','roa_pct','utilidad_acumulada_mmclp']
SQL_INSERT = 'INSERT INTO indicadores_bancarios ('+','.join(COLUMNAS)+') VALUES ('+','.join('?' for _ in COLUMNAS)+')'

@contextmanager
def conexion(root):
    carpeta = Path(root) / 'datos/procesados'
    carpeta.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(carpeta / 'bancos.sqlite', timeout=30)
    con.execute('CREATE TABLE IF NOT EXISTS indicadores_bancarios (fecha TEXT, banco TEXT, tipo_entidad TEXT, alcance TEXT, archivo_fuente TEXT, colocaciones_mmclp REAL, colocaciones_costo_amortizado_mmclp REAL, morosidad_90_pct REAL, provisiones_clasificacion_pct REAL, eficiencia_gastos_ingresos_pct REAL, UNIQUE(fecha,banco))')
    con.execute('CREATE TABLE IF NOT EXISTS trazabilidad (fecha TEXT, banco TEXT, indicador TEXT, valor REAL, unidad TEXT, archivo TEXT, hoja TEXT, celda TEXT)')
    con.execute('CREATE TABLE IF NOT EXISTS cargas (fecha TEXT, sha256 TEXT, archivo_original TEXT, archivo_guardado TEXT, guardado_en TEXT DEFAULT CURRENT_TIMESTAMP, UNIQUE(fecha, sha256))')
    # Migración aditiva: conserva los campos y registros de versiones anteriores.
    actuales={row[1] for row in con.execute('PRAGMA table_info(indicadores_bancarios)')}
    for campo in NUEVAS:
        if campo not in actuales:
            con.execute(f'ALTER TABLE indicadores_bancarios ADD COLUMN {campo} REAL')
    try:
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()

def leer(root):
    with conexion(root) as con:
        tabla = pd.read_sql_query('SELECT * FROM indicadores_bancarios ORDER BY fecha,banco', con)
        if tabla.empty:
            ruta = Path(root) / 'datos/procesados/bancos_cmf.csv'
            if ruta.exists():
                tabla = pd.read_csv(ruta)
                tabla=tabla.reindex(columns=COLUMNAS)
                con.executemany(SQL_INSERT,tabla.where(pd.notna(tabla),None).itertuples(index=False,name=None))
        if not tabla.empty:
            sistemas=tabla[tabla.tipo_entidad=='sistema'].set_index('fecha').colocaciones_mmclp
            if sistemas.index.duplicated().any():
                raise ValueError('Hay más de un total oficial del sistema para el mismo mes.')
            denominador=tabla.fecha.map(sistemas)
            tabla['participacion_colocaciones_pct']=tabla.colocaciones_mmclp/denominador.where(denominador>0)*100
        return tabla

def preparar(archivos):
    lote, meses = [], set()
    for nombre, contenido in archivos:
        nombre = Path(nombre.replace('\\', '/')).name
        tabla, traza = extraer(BytesIO(contenido), nombre)
        fecha = tabla.fecha.iloc[0]
        if fecha in meses:
            raise ValueError(f'Hay dos archivos para {fecha}. Selecciona solo uno por mes.')
        meses.add(fecha)
        lote.append((nombre, contenido, tabla, traza))
    return lote

def guardar(root, lote, reemplazar=False):
    if not lote:
        raise ValueError('No hay archivos para guardar.')
    root = Path(root)
    archivo_dir = root / 'datos/archivados'
    archivo_dir.mkdir(parents=True, exist_ok=True)
    existentes = leer(root)
    repetidos = sorted(set(existentes.fecha) & {item[2].fecha.iloc[0] for item in lote})
    if repetidos and not reemplazar:
        raise ValueError('Ya existen estos períodos: ' + ', '.join(repetidos) + '. Marca la opción de reemplazo para actualizarlos.')
    with conexion(root) as con:
        for nombre, contenido, tabla, traza in lote:
            fecha = tabla.fecha.iloc[0]
            sha = hashlib.sha256(contenido).hexdigest()
            nombre_seguro = f'{fecha}_{sha}.xlsx'
            destino = archivo_dir / nombre_seguro
            if not destino.exists():
                destino.write_bytes(contenido)
            con.execute('DELETE FROM indicadores_bancarios WHERE fecha=?', (fecha,))
            con.execute('DELETE FROM trazabilidad WHERE fecha=?', (fecha,))
            con.executemany(SQL_INSERT, tabla[COLUMNAS].itertuples(index=False, name=None))
            con.executemany('INSERT INTO trazabilidad VALUES (?,?,?,?,?,?,?,?)', traza[TRAZAS].itertuples(index=False, name=None))
            con.execute('INSERT OR IGNORE INTO cargas(fecha,sha256,archivo_original,archivo_guardado) VALUES (?,?,?,?)', (fecha,sha,nombre,nombre_seguro))
    return len(lote)

def comparar(tabla, actual, referencia, indicador, bancos):
    columnas = ['banco', indicador]
    a = tabla[(tabla.fecha == actual) & (tabla.tipo_entidad == 'banco') & tabla.banco.isin(bancos)][columnas].rename(columns={indicador:'actual'})
    b = tabla[(tabla.fecha == referencia) & (tabla.tipo_entidad == 'banco') & tabla.banco.isin(bancos)][columnas].rename(columns={indicador:'referencia'})
    resultado = a.merge(b, on='banco', how='outer', validate='one_to_one')
    resultado['cambio'] = resultado.actual - resultado.referencia
    if indicador.endswith('mmclp'):
        base=resultado.referencia.where(resultado.referencia>0) if indicador=='utilidad_acumulada_mmclp' else resultado.referencia.where(resultado.referencia!=0)
        resultado['variacion_pct'] = (resultado.actual / base - 1) * 100
    return resultado

def completar_historial(root):
    """Completa campos nuevos solo con archivos que coinciden con las cifras guardadas."""
    root=Path(root)
    tabla=leer(root)
    reporte=[]
    if tabla.empty: return reporte
    pendientes=tabla[tabla[NUEVAS].isna().any(axis=1)].fecha.unique()
    candidatos=list((root/'datos/originales').glob('*.xlsx'))+list((root/'datos/archivados').glob('*.xlsx'))
    anteriores=[c for c in MAPA if c not in NUEVAS]
    for fecha in pendientes:
        actuales=tabla[tabla.fecha==fecha].set_index('banco').sort_index()
        compatibles=[]
        vistos=set()
        for ruta in candidatos:
            try:
                contenido=ruta.read_bytes()
            except OSError:
                continue
            sha=hashlib.sha256(contenido).hexdigest()
            if sha in vistos: continue
            vistos.add(sha)
            try:
                nuevo,traza=extraer(BytesIO(contenido),actuales.archivo_fuente.iloc[0])
            except (ValueError,KeyError,OSError,zipfile.BadZipFile):
                continue
            if nuevo.fecha.iloc[0]!=fecha: continue
            nuevo=nuevo.set_index('banco').sort_index()
            if not actuales.index.equals(nuevo.index): continue
            iguales=True
            for col in anteriores:
                diferencia=(actuales[col]-nuevo[col]).abs()
                tolerancia=nuevo[col].abs()*1e-10+1e-8
                if not (diferencia<=tolerancia).all(): iguales=False; break
            if iguales: compatibles.append((nuevo,traza))
        if not compatibles:
            reporte.append(f'{fecha}: falta un Excel que coincida con las cifras guardadas. Vuelve a cargar el reporte de ese mes.')
            continue
        nuevo,traza=compatibles[0]
        if any(not nuevo[NUEVAS].equals(otro[NUEVAS]) for otro,_ in compatibles[1:]):
            reporte.append(f'{fecha}: hay revisiones distintas; vuelve a cargar la versión que deseas conservar.')
            continue
        with conexion(root) as con:
            for banco,fila in nuevo.iterrows():
                con.execute('UPDATE indicadores_bancarios SET roe_pct=COALESCE(roe_pct,?),roa_pct=COALESCE(roa_pct,?),utilidad_acumulada_mmclp=COALESCE(utilidad_acumulada_mmclp,?) WHERE fecha=? AND banco=?',(*[float(fila[c]) for c in NUEVAS],fecha,banco))
            for campo in NUEVAS:
                con.execute('DELETE FROM trazabilidad WHERE fecha=? AND indicador=?',(fecha,campo))
            filas=traza[traza.indicador.isin(NUEVAS)]
            con.executemany('INSERT INTO trazabilidad VALUES (?,?,?,?,?,?,?,?)',filas[TRAZAS].itertuples(index=False,name=None))
        reporte.append(f'{fecha}: ROE, ROA y utilidad incorporados sin cambiar las cifras anteriores.')
    return reporte