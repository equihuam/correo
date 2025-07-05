# Outlook 365

pacman::p_load(Microsoft365R, dplyr, stringr, progress)

#outl <- get_personal_outlook()

outlb <- get_business_outlook()

inbox_ref <- outlb$get_inbox()

correos <- inbox_ref$list_emails(by="from desc", 
                                 search = "RECIBO", 
                                 pagesize = 200)

tabla_correo <- tibble(fecha = character(), 
                       concepto = character(), 
                       archivo = character())

# Barra de avance
pb <- progress_bar$new(format = " Descargando [:bar] :percent eta: :eta",
                       total = length(correos))


for (correo in correos)
{
  remitente <-  correo$properties$sender$emailAddress$address
  asunto <- correo$properties$subject
  if (str_detect(remitente, "sistemagrp@") &
      str_detect(asunto, "NOMINA QNA "))
  {
      d  <- correo$list_attachments()
      recibo_xml_nombre <- d[[1]]$properties$name
      recibo_pdf_nombre <- d[[2]]$properties$name
      fecha  <- as.POSIXlt( as.Date(correo$properties$receivedDateTime))
      year <- fecha$year + 1900
      mes <- fecha$mon + 1
      d[[1]]$download(dest = paste0("recibos/", year, "/",
                                    mes, "-",
                                    year, "-",
                                    recibo_xml_nombre), 
                      overwrite = TRUE)
      d[[2]]$download(dest = paste0("recibos/", year, "/",
                                    mes, "-",
                                    year, "-",
                                    recibo_pdf_nombre), 
                      overwrite = TRUE)
      fecha_str <- correo$properties$receivedDateTime
      qna <- str_extract(asunto, "(?<=COS DE )(.*)")
      doc <- str_remove(recibo_xml_nombre, ".xml")
      tabla_correo <- bind_rows(tabla_correo, c(fecha = fecha_str,
                                                concepto = qna,
                                                archivo = doc))
      pb$tick()
  }
}




tabla_correo <- tibble(fecha = character(), 
                       concepto = character(), 
                       archivo = character())

for (correo in correos)
{
  remitente <-  correo$properties$sender$emailAddress$address
  asunto <- correo$properties$subject
  if (str_detect(remitente, "sistemagrp@") &
      str_detect(asunto, "NOMINA QNA "))
  {
    fecha <- correo$properties$receivedDateTime
    qna <- str_extract(asunto, "(?<=COS DE )(.*)")
    doc <- str_remove(recibo_pdf_nombre <- d[[1]]$properties$name, ".xml")
    tabla_correo <- bind_rows(tabla_correo, c(fecha = fecha,
                              concepto = qna,
                              archivo = doc))
    pb$tick()
    
  }
}




# Conciliae conda y r-env????
#Sys.setenv(RENV_PATHS_CACHE = '~/0 Versiones/lineas-rays/renv/cache')

#renv::use_python(type = 'conda', name = 'r-py')



