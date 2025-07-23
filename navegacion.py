import marimo

__generated_with = "0.14.11"
app = marimo.App(width="full")


@app.cell
def _():
    import marimo as mo
    return (mo,)


@app.cell
def menu(mo):
    mo.sidebar(
        [
            mo.md(f"<h4>Análisis de datos</h1>"),
            mo.nav_menu(
                {
                    "Sueldos": {
                        "/?file=compulab.py": "* Controles e informes",
                        "/?file=estadisticos.py": "* Estadísticas",
                        "/?file=pendientes.py": "* Advertencias",
                    },                
                },
                orientation="vertical"
            )
        ]
    )
    return


if __name__ == "__main__":
    app.run()
