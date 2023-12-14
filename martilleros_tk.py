import pandas as pd
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import messagebox, ttk


class Ventana:
    def __init__(self):
        self.ventana_principal = tk.Tk()
        self.ventana_principal.config(width=800, height=400)
        self.ventana_principal.resizable(False, False)
        self.ventana_principal.title(
            "Colegio de Martilleros de Dolores - by Mariano Francisco v0.5"
        )

        self.conectar = ttk.Button(
            self.ventana_principal,
            text="    --==Conectar==--    ",
            command=self.conectar_csv,
        )
        self.conectar.place(x=30, y=30)

        self.datos1 = tuple(range(2022, 2050))
        self.anios_num = tk.StringVar()
        self.anios = ttk.Combobox(
            self.ventana_principal, textvariable=self.anios_num, state="readonly", values=self.datos1, width=10
        )
        self.anios.current(2)
        self.anios.place(x=200, y=34)

        self.datos2 = tuple(range(1, 4))
        self.cuatri_num = tk.StringVar()
        self.cuatri = ttk.Combobox(
            self.ventana_principal, textvariable=self.cuatri_num, state="readonly", values=self.datos2, width=5
        )
        self.cuatri.current(0)
        self.cuatri.place(x=310, y=34)

        self.generar = ttk.Button(
            self.ventana_principal,
            text="    Generar   ",
            command=self.generar_mdf,
            state=tk.DISABLED,
        )
        self.generar.place(x=600, y=30)

        self.salir = ttk.Button(
            self.ventana_principal,
            text="    Salir   ",
            command=self.ventana_principal.quit,
        ).place(x=700, y=30)

        self.frame1 = tk.LabelFrame(self.ventana_principal, text="Datos de Excel")
        self.frame1.place(x=10, y=80, height=280, width=780)

        ## Treeview Widget
        self.tv1 = ttk.Treeview(self.frame1)
        self.tv1.place(relheight=1, relwidth=1)
        self.treescrolly = tk.Scrollbar(
            self.frame1, orient="vertical", command=self.tv1.yview
        )
        self.treescrollx = tk.Scrollbar(
            self.frame1, orient="horizontal", command=self.tv1.xview
        )
        self.tv1.configure(
            xscrollcommand=self.treescrollx.set, yscrollcommand=self.treescrolly.set
        )
        self.treescrollx.pack(side="bottom", fill="x")
        self.treescrolly.pack(side="right", fill="y")

        self.barra = ttk.Label(
            self.ventana_principal,
            text="Conversor para matriculados del Colegio de Martilleros de Dolores, BA",
            compound="bottom",
        )
        self.barra.place(x=10, y=370)

        self.ventana_principal.mainloop()

    def conectar_csv(self):
        if ruta.is_file():
            self.df = pd.read_excel(ruta)
            if self.df.shape[1] == 23:
                self.tv1["column"] = list(self.df.columns)
                self.tv1["show"] = "headings"
                for column in self.tv1["columns"]:
                    self.tv1.heading(column, text=column)
                self.df_rows = self.df.to_numpy().tolist()
                for row in self.df_rows:
                    self.tv1.insert("", "end", values=row)
                messagebox.showinfo(
                    message="Archivo OK - Padrón conectado",
                    title="Colegio de Martilleros",
                )
                self.conectar["state"] = tk.DISABLED
                self.generar["state"] = tk.NORMAL
            else:
                messagebox.showinfo(
                    message="--==Padrón malformado==--", title="Error de columnas"
                )
        else:
            messagebox.showinfo(
                message=f"--==El archivo {ruta} no existe!==--",
                title="Error de archivo",
            )

    def generar_mdf(self):
        """Genera el archivo de salida"""

        def partidos(p: str) -> str:
            if p.lower() not in PARTIDO:
                return "0002"
            if p.lower() == PARTIDO[0]:
                return "0001"
            else:
                return "0003"

        def provincias1(fila: dict) -> str:
            if fila["Provincia_residencia"].lower() not in PARTIDO:
                return "0002"
            return str(PROVINCIA.index(fila["Provincia_residencia"].upper()) + 1).zfill(
                4
            )

        def provincias2(fila: dict) -> str:
            if fila["Provincia_actividad"].lower() not in PARTIDO:
                return "0002"
            return str(PROVINCIA.index(fila["Provincia_actividad"].upper()) + 1).zfill(
                4
            )

        IDENTIFICACION = (
            "CUIT",
            "CUIL",
            "CDI",
            "DNI",
            "CPF",
            "DNIE",
            "LC",
            "LE",
            "PASAPORTE",
        )
        PARTIDO = ("capital federal", "otras provincias")
        PROVINCIA = (
            "CAPITAL FEDERAL",
            "BUENOS AIRES",
            "CATAMARCA",
            "CORDOBA",
            "CORRIENTES",
            "CHACO",
        )

        if ruta.is_file():
            self.df = pd.read_excel(ruta)
        else:
            return None

        if self.df.shape[1] != 23:
            return None

        # Datos
        self.fecha = str(date.today()).replace("-", "")
        self.limite = ("0430", "0831", "1231")
        # for id, lim in enumerate(self.limite, start=1):
        #     if self.fecha[4:] <= lim:
        #         trimestre = id
        #         break
        self.periodo = f"{self.anios_num.get()}{self.cuatri_num.get()}"
        #print(self.periodo)
        #self.periodo = f"{self.fecha[:4]}{trimestre}"
        #self.periodo = "20233"

        # remove special character
        self.df.columns = self.df.columns.str.replace(" ", "_")
        if self.fecha < str(max(self.df["Fecha_de_matriculación_AAAAMMDD"])):
            messagebox.showinfo(
                message=f"Hay matriculados fuera del período",
                title="Error de fechas",
            )
            return None

        # Cambia los NaN por cadena vacía
        self.df = self.df.fillna("")

        # Inserta las dos filas del inicio
        self.df.insert(0, "DDJJ_tipo", "E")
        self.df.insert(1, "DDJJ_periodo", self.periodo)

        # Inserta las cols extras en vacío
        cols = {
            10: "Sin_Numero_DomRes",
            13: "Torre_DomRes",
            14: "Piso_DomRes",
            15: "Dto_DomRes",
            16: "Mzna_DomRes",
            21: "PreT_DomRes",
            22: "Tel_DomRes",
            23: "PreF_DomRes",
            24: "Fax_DomRes",
            25: "Mail_DomRes",
            29: "Sin_Numero_DomAct",
            32: "Torre_DomAct",
            33: "Piso_DomAct",
            34: "Dto_DomAct",
            35: "Mzna_DomAct",
            40: "PreT_DomAct",
            41: "Tel_DomAct",
            42: "PreF_DomAct",
            43: "Fax_DomAct",
            44: "Mail_DomAct",
            46: "SinValor",
        }
        for c, v in cols.items():
            self.df.insert(c - 1, v, "")

        # Saca decimales o pasa a STR y remueve espacios con strip()
        convertir = ("N°_residencia", "N°_actividad")
        for col in self.df.columns:
            if col in convertir:
                self.df[col] = self.df[col].apply(str).str[:-2].str.strip()
            else:
                self.df[col] = self.df[col].apply(str).str.strip()

        # Si es anterior a 01/01/2004 lo actualiza
        self.df.loc[
            self.df["Fecha_de_matriculación_AAAAMMDD"] < "20040101",
            "Fecha_de_matriculación_AAAAMMDD",
        ] = "20040101"

        # Cambia del tipo de DOC a código IDENTIFICACION
        self.df["Tipo_DOC"] = (
            self.df["Tipo_DOC"].map(lambda x: IDENTIFICACION.index(x) + 1).apply(str)
        )

        # Sección Domicilio de Residencia

        # Crea col y asigna con nro de residencia 0 y sin nro 1
        self.df["Sin_Numero_DomRes"] = (
            self.df["N°_residencia"].map(lambda x: int(not (x))).apply(str)
        )

        # Cambia las cols 18 y 19 calculadas según la Localidad_residencia
        self.df["Partido_residencia"] = self.df["Provincia_residencia"].map(partidos)
        self.df["Provincia_residencia"] = self.df.apply(provincias1, axis=1)
        self.df["Localidad_residencia"] = self.df["Localidad_residencia"].str.upper()

        # En caso de no informar dom residencia
        self.df.loc[
            self.df["N°_residencia"] == "", "Observaciones_Dom_residencia"
        ] = "No informa"

        # Sección Domicilio de Actividad Profesional

        # Crea col y asigna con nro de residencia 0 y sin nro 1
        self.df["Sin_Numero_DomAct"] = (
            self.df["N°_actividad"].map(lambda x: int(not (x))).apply(str)
        )

        # Cambia las cols 37 y 38 calculadas según la Localidad_residencia
        self.df.loc[
            self.df["Provincia_actividad"] == "Capital Federal", "Localidad_actividad"
        ] = "Capital Federal"
        self.df["Partido_actividad"] = self.df["Provincia_actividad"].map(partidos)
        self.df["Provincia_actividad"] = self.df.apply(provincias2, axis=1)
        self.df["Localidad_actividad"] = self.df["Localidad_actividad"].str.upper()

        # En caso de no informar dom residencia
        self.df.loc[self.df["N°_actividad"] == "", "Observaciones"] = "No informa"

        # Genera el archivo para importar en ARBA
        archivo = Path(f"./padron-{self.periodo}-{self.fecha}.mdf")
        self.df.to_csv(
            archivo,
            sep=";",
            index=False,
            header=False,
            encoding="latin1",
            lineterminator="\r\n",
        )
        messagebox.showinfo(
            message=f"Se generó el archivo {archivo}",
            title="Archivo generado",
        )
        return None


if __name__ == "__main__":
    ruta = Path("./padron2.xls")
    app = Ventana()
