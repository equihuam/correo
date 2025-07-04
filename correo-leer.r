     l
# GMAIL

library(gmailr)

gm_auth_configure(path = "_creds/client_secret.json")

# Authenticate with the new cache, store tokens in .secret
gm_auth(cache = ".secret")

# view the latest thread
conv_hilo <- gm_threads(search = "recibo", num_results = 15, )


# retrieve the latest thread by retrieving the first ID

ultimo_hilo <- gm_thread(gm_id(conv_hilo)[[1]])

# The messages in the thread will now be in a list
ultimo_hilo$messages

# Retrieve parts of a specific message with the accessors
mi_msg <- ultimo_hilo$messages[[1]]


gm_to(mi_msg) 

gm_from(mi_msg)

gm_date(mi_msg)

gm_subject(mi_msg)

gm_body(mi_msg)

# If a message has attachments, download them all locally with `gm_save_attachments()`.
gm_save_attachments(mi_msg, path = "_adjuntos/")



# Outlook 365


pacman::p_load(Microsoft365R, dplyr, stringr)

#outl <- get_personal_outlook()

outlb <- get_business_outlook()

inbox_ref <- outlb$get_inbox()

inbox_ref$get_list_pager(lst = )

correos <- inbox_ref$list_emails(by="from desc", search = "RECIBO", pagesize = 200)



# Conciliae conda y r-env????
#Sys.setenv(RENV_PATHS_CACHE = '~/0 Versiones/lineas-rays/renv/cache')

#renv::use_python(type = 'conda', name = 'r-py')



 