import base64
import pandas as pd
import plotly.express as px
import os
import zipfile


def dir_to_zip(dir_path):
    if not os.path.isdir(dir_path):
        raise ValueError(f"The provided path {dir_path} is not a valid directory.")
    parent_dir = os.path.abspath(os.path.join(dir_path, os.pardir))
    dir_name = os.path.basename(os.path.normpath(dir_path))
    zip_filename = os.path.join(parent_dir, f"{dir_name}.zip")
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(dir_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, dir_path))



def diagram(title, sheet_name, filepath, nrows, skiprows):
    df = pd.read_excel(
        filepath,
        sheet_name=sheet_name,
        skiprows=skiprows,
        nrows=nrows
    )
    df_long = df.rename(
        columns={'Unnamed: 0': 'Témakör'}).melt(id_vars='Témakör', var_name='Év', value_name='Összeg')
    df_long['Összeg'] = df_long['Összeg'] * 1000
    df_long['Összeg_millió'] = df_long['Összeg'] / 1e6

    fig = px.bar(
        df_long,
        x='Év',
        y='Összeg_millió',
        color='Témakör',
        title=title,
        labels={'Összeg_millió': 'Összeg (millió Ft)'},
        barmode='stack',
        text="Témakör",
        color_discrete_sequence=[
            "#008345",
            "#f7c900",
            "#ae1c2f",
            "#0053a2",
            "#e9802d"
        ],
    )

    fig.update_layout(
        barcornerradius=10,
        yaxis=dict(
            tickformat=',.0f',
            title='Összeg (millió Ft)'
        ),
        bargap=0.15,
        bargroupgap=0.1,
        title={'text': '', },
        xaxis=dict(
            tickformat='.0f',
            tickmode='array',
            tickvals=sorted(df_long['Év'].unique()),
            ticktext=[str(year) for year in sorted(df_long['Év'].unique())],
            title='Év'
        ),
        legend=dict(
            title='',
            orientation='h',
            x=0.5,
            y=1.1,
            xanchor='center',
            yanchor='top',
            valign='middle'
        ),
        margin=dict(t=100, l=20, r=20, b=20),
    )

    fig.update_yaxes(tickprefix="")
    fig.update_xaxes(tickprefix="")

    with open("resource/img.png", "rb") as image_file:
        encoded_image = base64.b64encode(image_file.read()).decode()

    fig.add_layout_image(
        dict(
            source=f"data:image/png;base64,{encoded_image}",
            xref="paper", yref="paper",
            x=0.11, y=0.98,
            sizex=0.2, sizey=0.2,
            xanchor="center", yanchor="top"
        )
    )

    fig.show()

    os.makedirs("diagrammok", exist_ok=True)

    with open(f"diagrammok/{title}.svg", "wb") as file:
        file.write(fig.to_image("svg", height=620, width=1020))

    with open(f"diagrammok/{title}.html", "w") as file:
        print()
        file.write(fig.to_html(full_html=True))


diagram("Kiadások", "Kiadások", "resource/gazd.xlsx", 5, 9)
diagram("Bevételek", "Bevételek", "resource/gazd.xlsx", 4, 8)
dir_to_zip("diagrammok")
