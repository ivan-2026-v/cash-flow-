# Cash flow Nexiu — automatizacion

Reconstruye el cash flow mensual de Nexiu desde los movimientos bancarios crudos,
replicando la logica del archivo que Ivan mantenia a mano
(`Cash_Flow_report_2608.xlsx`).

```
python3 src/cashflow.py 2026-01               # genera las dos vistas
python3 src/cashflow.py 2026-01 --base caja   # solo una
```

Salida en `output/`: un resumen en markdown y el detalle movimiento por
movimiento en csv, para poder auditar cualquier numero hasta su fila de origen.

## Las dos vistas

| Vista | Mes de cada movimiento | Para que sirve |
|---|---|---|
| `caja` | fecha real del banco | **Concilia exacto contra el extracto.** Es la verdad del saldo. |
| `devengo` | mes al que corresponde el gasto | Es la vista del reporte manual de Ivan. Sueldos de fundadores y renta se pagan al mes siguiente y se reasignan hacia atras. |

## Como se clasifica un movimiento

1. **Normalizacion.** Cada fuente se lee con su parser (`mercury_xlsx`,
   `dolarapp_csv`) y se lleva a un formato comun en USD. En DolarApp el importe
   en USD es `base_amount`, ya convertido al fx del dia por el banco.
2. **Diccionario** (`config/dictionary.csv`, 167 entradas exportadas del archivo
   de Ivan). Busqueda exacta por merchant — en Mercury contra `Description`, en
   DolarApp contra `name`, igual que los `VLOOKUP` del original. Devuelve
   `Statement` (P&L / BS / Intercompany), `Type 2` y `Type 3`.
   La comparacion ignora acentos y mayusculas, porque el archivo original tiene
   el mismo merchant escrito de varias formas.
3. **Ajustes** (`config/mapping.yaml`). Reglas transversales que pisan al
   diccionario:
   - Facebook con la tarjeta terminada en **2354** → `3. Platform / 3. Messaging`
   - Suscripcion **Anthropic de 200 USD** → `6. Tech salaries / 1. Salaries`
     (los cargos de Anthropic por otro importe siguen siendo Platform)
   - Plata que **entra** de Ivan (Wise o transferencia) → `BS / 0. Investments`,
     porque es aporte de socio y no ingreso. La misma contraparte cuando **sale**
     es sueldo de fundador: la regla depende del signo.
4. **Correcciones fila por fila** (`config/overrides.csv`). Lo que Ivan hacia
   editando celdas sueltas: reasignar el mes y/o pisar la clasificacion de un
   movimiento puntual. Se indexan por cuenta + fecha + merchant + importe, asi
   una correccion nunca se aplica por error a un movimiento parecido de otro mes.
5. **Filtro de estado.** Solo entran los `Sent` / `COMPLETED`. Los `FAILED` y
   `REVERTED` se leen y quedan en el detalle, pero no suman.
6. **Agregacion** por `Type 2`, con `Type 2_B` como segundo nivel
   (`6. Tech salaries`, `7. Founders` y `8. Other fixed costs` colapsan en
   `6. Other fixed and admin`).

Solo lo marcado `P&L` entra al cash flow operativo. `BS` e `Intercompany` mueven
caja pero se muestran aparte, y por eso la conciliacion cierra igual.

## Conciliacion

El saldo inicial y final de cada cuenta sale de los PDF en
`data/bank-statements` y esta declarado en `config/accounts.yaml`. El script
compara `saldo inicial + neto calculado` contra el saldo final del extracto.
**Esa diferencia tiene que dar cero**; si no da, hay un movimiento que falta o
que sobra.

## Estructura

```
config/
  accounts.yaml      cuentas, archivos por mes y saldos de extracto
  dictionary.csv     merchant -> Statement / Type 2 / Type 3
  mapping.yaml       estados, Type 2_B, orden del reporte y ajustes transversales
  overrides.csv      correcciones manuales fila por fila
data/
  raw-data/          exports de movimientos
  bank-statements/   PDF de extractos (fuente de la verdad de los saldos)
src/cashflow.py
output/
```

## Agregar un mes

1. Dejar los exports en `data/raw-data/` y los PDF en `data/bank-statements/`.
2. Sumar el mes a `files` de cada cuenta en `config/accounts.yaml`, y los saldos
   de apertura y cierre a `statement_balances`.
3. Correr el script y revisar dos cosas: que la conciliacion de la vista `caja`
   de cero, y que no queden avisos de `SIN DICCIONARIO`. Un merchant nuevo se
   agrega a `config/dictionary.csv`.
