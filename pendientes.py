import marimo

__generated_with = "0.14.11"
app = marimo.App(width="full")


@app.cell
def _():
    import marimo as mo
    return (mo,)


@app.cell
def menu(mo):
    mo.md(
        """
    Ciertas funcionalidades no se encuentran disponibles, pendientes de ser implementadas a futuro, a saber:

    * <b>Descuentos de obra social sobre conceptos no remunerativos: </b>Los casos en que a los empleados les pueda corresponder un descuento de obra social sobre conceptos no remunerativos, aún no se encuentran contemplados, será necesario verificar manualmente si los importes descontados son los correctos.
    * <b>Familiares adherentes a la obra social: </b>Los casos en que los empleados cuenten con adherentes a la obra social y por lo tanto les corresponda un descuento mayor al habitual, aún no se encuentran contemplados, será necesario verificar manualmente si los importes descontados son los correctos.
    * <b>Control de sindicatos: </b>De momento no existe control alguno sobre descuentos sindicales, por lo tanto, deberán ser calculados manualmente de manera integral.
    """
    )
    return


@app.cell
async def _():
    # Carga la barra lateral de navegación
    from navegacion import app

    sidebar = await app.embed()

    sidebar.output
    return


if __name__ == "__main__":
    app.run()
