# Genera emails con base en Outlook 365

pacman::p_load(Microsoft365R, dplyr, stringr, blastula, readr, gt)

#outl <- get_personal_outlook()

# Prepara la interacción con Outlook con "cuenta profesional"
outlb <- get_business_outlook()

# Datos base
datos_envio    <- read.csv("tareas/contratos-ciab/dat-contactos.csv")
datos_act      <- read.csv("tareas/contratos-ciab/dat-conceptos.csv")
datos_prods    <- read.csv("tareas/contratos-ciab/dat-entregables.csv")
mensaje_base   <- read_file("tareas/contratos-ciab/mensaje-base.md")
img_firma_html <- add_image(file ="tareas/contratos-ciab/firma-Miguel-2025.png",
                            width =600)
anexo <- "tareas/contratos-ciab/carta_datos_bancarios.docx"
lista_cc <- source(file = "tareas/contratos-ciab/dat-copias.txt", )$value


# Genera los mensajes personalizados  
for (i in (15:length(datos_envio$nombre)))
{
  if (!is.na(datos_envio$email[i]))
  {
    nombre <-  datos_envio$nombre[i]
    saludo <-  datos_envio$saludo[i]
    email <- datos_envio$email[i]
    act <- datos_envio$actividad[i]
    concepto <- datos_act$Concepto[datos_act$Actividad == act]
    
    tabla <- datos_prods |> 
      filter(Actividad == act) |>
      select(sec, Concepto, Entregable) |> 
      gt() |>
      cols_width(
        sec ~ px(20),
        Concepto ~ px(350),
        Entregable ~ px(150)) |> 
      tab_style(
        style = list(cell_borders(sides = c("l", "r", "t", "b"),
                                  color = "gray", weight = px(1))),
        locations = list(cells_column_labels(), cells_body())) |> 
      opt_table_font(size = "10pt") |> 
      opt_stylize(style = 6, color = 'gray')

    mensaje_md <- (md(c(str_glue(mensaje_base, img_firma_html)))) 
    
    bl_em <- compose_email(
      body=mensaje_md,
      footer=md("enviado via Microsoft365R"))
    em <- outlb$create_email(bl_em, 
                             subject="Solicitud de cotización", 
                             to = email,
                             cc = lista_cc)
    
    # add an attachment and send it
    em$add_attachment(anexo)
  }

}



# Genera mensajes personalizados solicitando dato faltante  
mensaje_adicion <-  read_file("tareas/contratos-ciab/mensaje-dato-faltante.md")

for (i in (15:length(datos_envio$nombre)))
{
  if (!is.na(datos_envio$email[i]))
  {
    nombre <-  datos_envio$nombre[i]
    saludo <-  datos_envio$saludo[i]
    email <- datos_envio$email[i]

    mensaje_md <- (md(c(str_glue(mensaje_adicion, img_firma_html)))) 
    
    bl_em <- compose_email(
      body=mensaje_md,
      footer=md("enviado via Microsoft365R"))
    em <- outlb$create_email(bl_em, 
                             subject="Solicitud de cotización", 
                             to = email,
                             cc = lista_cc)
  }
}  


#em$send() 

