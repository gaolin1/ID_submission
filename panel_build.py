import panel as pn
import pandas as pd

pn.extension()

file_input = pn.widgets.FileInput(accept='.xlsx')
#df = pn.widgets.DataFrame(pd.read_excel(file_input.value))

layout = pn.Column(
    pn.Row(file_input),
    pn.Row(file_input)
)

pn.template.FastListTemplate(
    site="Panel", title="Example", main=[layout],
).servable()