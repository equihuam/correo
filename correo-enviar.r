# Outlook 365

pacman::p_load(Microsoft365R, dplyr, stringr, blastula, readr, rmarkdown)

#outl <- get_personal_outlook()
outlb <- get_business_outlook()

bl_mensaje <- read_file("mensaje.md")

datos_envio <-  read.csv("mensajes-datos.csv")

for (i in (1:length(datos_envio)))
{
  
  if (!is.na(datos_envio$email[i]))
  {
    nombre <-  datos_envio$nombre[i]
    asunto <-  datos_envio$asunto[i]
    mensaje_md <- (md(str_glue(bl_mensaje)))
  
    bl_em <- compose_email(
      body=mensaje_md,
      footer=md("enviado via Microsoft365R"))
    em <- outlb$create_email(bl_em, subject="Hola dede R", to="equihuam@gmail.com")
    
    # add an attachment and send it
    em$add_attachment("correo-leer.r")
  }

}

#em$send() 

