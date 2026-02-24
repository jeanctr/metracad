;;; ============================================================================
;;; METRACAD | Utilidad Automatizada de Medición de Longitudes en CAD
;;; Versión: 1.1.0
;;; Autor:   Jean Carlos
;;; Licencia: MIT
;;; ============================================================================
;;;
;;; COMANDOS DISPONIBLES:
;;;   METRACAD        - Mide longitudes de entidades seleccionadas
;;;   METRACADEXPORT  - Exporta el historial completo a CSV
;;;   METRACADCLIP    - Copia el último reporte al portapapeles
;;;
;;; HISTORIAL DE CAMBIOS:
;;;   1.0.0 - Versión inicial
;;;   1.0.1 - Corrección de variables locales, manejo de errores por entidad,
;;;           reemplazo de (exit) por flag abort
;;;   1.1.0 - Timestamp capturado una sola vez por sesión, comando METRACADCLIP,
;;;           opción de copiar al portapapeles al finalizar medición,
;;;           CSV enriquecido con detalle por capa/color
;;; ============================================================================

(vl-load-com)

;;; ============================================================================
;;; CONFIGURACIÓN GLOBAL
;;; ============================================================================

(setq *MC-VERSION*  "1.1.0")
(setq *MC-LAST*     nil)   ; Último reporte generado en texto plano
(setq *MC-HISTORY*  nil)   ; Historial de sesiones para exportación CSV
(setq *MC-DECIMALS* 2)     ; Precisión decimal por defecto
(setq *MC-UNIT*     "m")   ; Unidad de medida por defecto

;;; ============================================================================
;;; FUNCIONES UTILITARIAS
;;; ============================================================================

; Formatea un número real según la precisión decimal configurada globalmente
(defun MC:fmt (n) (rtos n 2 *MC-DECIMALS*))

; Rellena una cadena con espacios a la DERECHA hasta alcanzar el ancho indicado.
; Si la cadena es más larga que width, la trunca.
(defun MC:pad-r (str width)
  (if (> width (strlen str))
    (strcat str (substr "                                " 1 (- width (strlen str))))
    (substr str 1 width)))

; Rellena una cadena con espacios a la IZQUIERDA hasta alcanzar el ancho indicado.
; Si la cadena es más larga que width, la trunca.
(defun MC:pad-l (str width)
  (if (> width (strlen str))
    (strcat (substr "                                " 1 (- width (strlen str))) str)
    str))

; Obtiene la fecha y hora actual formateada desde la variable de sistema DATE.
; NOTA: Depende del menú de AutoCAD activo. Si retorna vacío, el menú no está cargado.
(defun MC:timestamp ()
  (menucmd "m=$(edtime,$(getvar,date),DD/MM/YYYY HH:MM:SS)"))

; Obtiene el nombre del archivo DWG activo mediante ActiveX
(defun MC:dwg-name ()
  (vla-get-Name (vla-get-ActiveDocument (vlax-get-acad-object))))

;;; Obtiene la longitud de una curva de forma segura.
;;; Encapsula la llamada en vl-catch-all-apply para evitar que errores
;;; en entidades problemáticas (LWPOLYLINE en bloques, splines abiertas, etc.)
;;; interrumpan el procesamiento del lote completo.
;;; Retorna 0.0 si la longitud no puede calcularse.
(defun MC:get-length (obj / len)
  (setq len nil)
  (vl-catch-all-apply
    '(lambda ()
       (setq len (vlax-curve-getDistAtParam
                   obj
                   (vlax-curve-getEndParam obj)))))
  (if (and len (numberp len) (> len 0)) len 0.0))

;;; Copia texto al portapapeles de Windows mediante archivo temporal.
;;; Escribe el contenido a un .tmp y lo pasa al comando nativo "clip" de Windows.
;;; Retorna T si el proceso se completó sin errores, nil si falló.
(defun MC:clipboard (text / tmp f result)
  (setq tmp    (strcat (getvar "TEMPPREFIX") "mc_clip.tmp")
        result nil)
  (vl-catch-all-apply
    '(lambda ()
       (setq f (open tmp "w"))
       (write-line text f)
       (close f)
       (vl-cmdf "_.SHELL"
                (strcat "cmd /c type \"" tmp "\" | clip"))
       (vl-file-delete tmp)
       (setq result T)))
  result)

;;; Escapa comas y comillas en una celda para formato CSV válido.
;;; Si el valor contiene comas, saltos de línea o comillas, lo encierra entre
;;; comillas dobles y escapa las comillas internas duplicándolas.
(defun MC:csv-cell (str / needs-quotes)
  (setq needs-quotes (or (vl-string-search "," str)
                         (vl-string-search "\"" str)
                         (vl-string-search "\n" str)))
  (if needs-quotes
    (strcat "\"" (vl-string-translate "\"" "\"\"" str) "\"")
    str))

;;; Construye una fila CSV a partir de una lista de valores string.
;;; Aplica escape a cada celda y las une con comas.
(defun MC:csv-row (cells)
  (apply 'strcat
    (cons (MC:csv-cell (car cells))
          (mapcar '(lambda (c) (strcat "," (MC:csv-cell c)))
                  (cdr cells)))))

;;; ============================================================================
;;; COMANDO PRINCIPAL: METRACAD
;;; Mide longitudes de entidades agrupadas por tipo y capa.
;;; Soporta tres modos: selección manual, filtro por capa, filtro por color.
;;; ============================================================================
(defun c:metracad (
  / *error* ss i ent obj type len lay entry report
    sum-lin sum-pol sum-arc sum-cir sum-spl sum-ell
    cnt-lin cnt-pol cnt-arc cnt-cir cnt-spl cnt-ell
    total-sum layer-data mode dec-opt unit-opt abort
    ts dwg-n clip-opt filter-ref filter-label
  )

  ; Manejador de errores local: muestra el mensaje solo si no es cancelación normal
  (defun *error* (msg)
    (if (not (member msg '("Function cancelled" "quit / exit abort")))
      (princ (strcat "\nError: " msg)))
    (princ))

  ; Captura timestamp y nombre de archivo UNA SOLA VEZ al inicio de la sesión.
  ; Esto garantiza consistencia entre el reporte visual y el registro en historial.
  (setq ts    (MC:timestamp)
        dwg-n (MC:dwg-name))

  ;; --- CONFIGURACIÓN DE SESIÓN ---

  ; Solicita la precisión decimal (2, 3 o 4 lugares decimales)
  (initget "2 3 4")
  (setq dec-opt (getkword (strcat "\nPrecisión [2/3/4] <" (itoa *MC-DECIMALS*) ">: ")))
  (if dec-opt (setq *MC-DECIMALS* (atoi dec-opt)))

  ; Solicita la unidad de medida que aparecerá en el reporte
  (initget "m cm mm ft")
  (setq unit-opt (getkword (strcat "\nUnidad [m/cm/mm/ft] <" *MC-UNIT* ">: ")))
  (if unit-opt (setq *MC-UNIT* unit-opt))

  ;; --- MODO DE SELECCIÓN ---

  ; Permite elegir entre selección manual, filtro por capa o filtro por color.
  ; El modo elegido también se registra en el historial para trazabilidad en CSV.
  (initget "Selection Layer Color")
  (setq mode (getkword "\nModo [Selection/Layer/Color] <Selection>: "))
  (setq mode (if mode mode "Selection"))

  ; Filtro de tipos de entidad soportados por el motor de medición
  (setq filter '((0 . "LINE,LWPOLYLINE,POLYLINE,ARC,CIRCLE,SPLINE,ELLIPSE")))

  ; filter-label almacena la descripción del filtro aplicado para el CSV.
  ; En modo Selection se llena después de la selección; en Layer/Color se
  ; obtiene del objeto de referencia elegido por el usuario.
  (setq filter-label "")

  (cond
    ; MODO SELECCIÓN: el usuario elige entidades directamente
    ((= mode "Selection")
     (setq ss (ssget filter))
     (setq filter-label "Selección manual"))

    ; MODO CAPA: selecciona todas las entidades del tipo correcto en la misma capa
    ((= mode "Layer")
     (if (setq filter-ref (car (entsel "\nSeleccione objeto para filtrar por capa: ")))
       (progn
         (setq lay (cdr (assoc 8 (entget filter-ref))))
         (setq filter-label (strcat "Capa: " lay))
         (setq ss (ssget "_X" (list (assoc 8 (entget filter-ref)) (car filter)))))))

    ; MODO COLOR: selecciona todas las entidades del tipo correcto con el mismo color
    ((= mode "Color")
     (if (setq filter-ref (car (entsel "\nSeleccione objeto para filtrar por color: ")))
       (progn
         (setq filter-label (strcat "Color: "
                               (itoa (cond ((cdr (assoc 62 (entget filter-ref)))) (256)))))
         (setq ss (ssget "_X" (list (cons 62
                                      (cond ((cdr (assoc 62 (entget filter-ref)))) (256)))
                                    (car filter))))))))

  ;; --- VALIDACIÓN ---

  ; Usa flag abort en lugar de (exit) para no interferir con el *error* local
  (setq abort nil)
  (if (null ss)
    (progn
      (princ "\nSelección vacía.")
      (setq abort T)))

  ;; --- PROCESAMIENTO ---

  (if (not abort)
    (progn
      ; Inicializa contadores y acumuladores por tipo de entidad
      (setq cnt-lin 0   cnt-pol 0   cnt-arc 0   cnt-cir 0   cnt-spl 0   cnt-ell 0
            sum-lin 0.0 sum-pol 0.0 sum-arc 0.0 sum-cir 0.0 sum-spl 0.0 sum-ell 0.0
            layer-data nil)

      ; Itera sobre cada entidad de la selección en orden inverso
      (repeat (setq i (sslength ss))
        (setq ent  (ssname ss (setq i (1- i)))
              obj  (vlax-ename->vla-object ent)
              type (cdr (assoc 0 (entget ent)))   ; Tipo: LINE, ARC, CIRCLE, etc.
              lay  (cdr (assoc 8 (entget ent)))   ; Nombre de capa de esta entidad
              len  (MC:get-length obj))            ; Longitud calculada con manejo de error

        ; Acumula por tipo de entidad
        (cond
          ((= type "LINE")
           (setq cnt-lin (1+ cnt-lin) sum-lin (+ sum-lin len)))
          ((wcmatch type "*POLYLINE")           ; Captura tanto POLYLINE como LWPOLYLINE
           (setq cnt-pol (1+ cnt-pol) sum-pol (+ sum-pol len)))
          ((= type "ARC")
           (setq cnt-arc (1+ cnt-arc) sum-arc (+ sum-arc len)))
          ((= type "CIRCLE")
           (setq cnt-cir (1+ cnt-cir) sum-cir (+ sum-cir len)))
          ((= type "SPLINE")
           (setq cnt-spl (1+ cnt-spl) sum-spl (+ sum-spl len)))
          ((= type "ELLIPSE")
           (setq cnt-ell (1+ cnt-ell) sum-ell (+ sum-ell len))))

        ; Acumula por capa para el detalle por capa en el CSV.
        ; layer-data es una lista de (nombre-capa longitud-total cantidad)
        (setq entry (assoc lay layer-data))
        (if entry
          ; Capa ya registrada: actualiza acumulados
          (setq layer-data (subst
                             (list lay (+ (cadr entry) len) (1+ (caddr entry)))
                             entry layer-data))
          ; Capa nueva: agrega entrada inicial
          (setq layer-data (append layer-data (list (list lay len 1))))))

      ; Longitud total de todas las entidades procesadas
      (setq total-sum (+ sum-lin sum-pol sum-arc sum-cir sum-spl sum-ell))

      ;; --- REPORTE EN PANTALLA ---

      ; Encabezado con metadatos de la sesión
      (setq report (strcat
        "\nMETRACAD v" *MC-VERSION* "\n"
        "Archivo: " dwg-n " | " ts "\n"
        "Modo: " mode " | Filtro: " filter-label " | Unidad: " *MC-UNIT* "\n"
        "------------------------------------------------\n"
        "TIPO          CANT        LONGITUD\n"
        "------------------------------------------------\n"))

      ; Función auxiliar: agrega fila al reporte solo si hay datos del tipo
      (defun add-row (lbl c s)
        (if (> c 0)
          (setq report (strcat report
            (MC:pad-r lbl 14)
            (MC:pad-l (itoa c) 5)
            (MC:pad-l (MC:fmt s) 15)
            " " *MC-UNIT* "\n"))))

      (add-row "Líneas"     cnt-lin sum-lin)
      (add-row "Polilíneas" cnt-pol sum-pol)
      (add-row "Arcos"      cnt-arc sum-arc)
      (add-row "Círculos"   cnt-cir sum-cir)
      (add-row "Splines"    cnt-spl sum-spl)
      (add-row "Elipses"    cnt-ell sum-ell)

      ; Fila de totales
      (setq report (strcat report
        "------------------------------------------------\n"
        (MC:pad-r "TOTAL" 14)
        (MC:pad-l (itoa (sslength ss)) 5)
        (MC:pad-l (MC:fmt total-sum) 15)
        " " *MC-UNIT* "\n"))

      ; Sección de detalle por capa (solo en reporte de pantalla)
      (if layer-data
        (progn
          (setq report (strcat report
            "\nDETALLE POR CAPA:\n"
            "------------------------------------------------\n"
            "CAPA          CANT        LONGITUD\n"
            "------------------------------------------------\n"))
          (foreach ld layer-data
            (setq report (strcat report
              (MC:pad-r (car ld) 14)
              (MC:pad-l (itoa (caddr ld)) 5)
              (MC:pad-l (MC:fmt (cadr ld)) 15)
              " " *MC-UNIT* "\n")))))

      ;; --- PERSISTENCIA ---

      ; Guarda el reporte de texto para METRACADCLIP
      (setq *MC-LAST* report)

      ; Registra la sesión en el historial para exportación CSV.
      ; Se guarda el detalle por capa como lista para generar filas individuales en CSV.
      (setq *MC-HISTORY* (append *MC-HISTORY*
        (list (list
          ts                        ; Fecha/hora
          dwg-n                     ; Nombre del archivo DWG
          mode                      ; Modo de selección usado
          filter-label              ; Descripción del filtro (capa, color, o manual)
          (itoa (sslength ss))      ; Cantidad total de entidades
          (MC:fmt total-sum)        ; Longitud total
          *MC-UNIT*                 ; Unidad de medida
          layer-data))))            ; Detalle por capa para CSV enriquecido

      ;; --- SALIDA ---

      (princ report)
      (alert report)

      ; Ofrece copiar el resultado al portapapeles al terminar
      (initget "Si No")
      (setq clip-opt (getkword "\n¿Copiar resultado al portapapeles? [Si/No] <No>: "))
      (if (= clip-opt "Si")
        (if (MC:clipboard report)
          (princ "\nResultado copiado al portapapeles.")
          (princ "\nNo se pudo copiar al portapapeles.")))))

  (princ))

;;; ============================================================================
;;; COMANDO: METRACADEXPORT
;;; Exporta el historial completo a un archivo CSV con detalle por capa.
;;;
;;; ESTRUCTURA DEL CSV:
;;;   El CSV genera DOS tipos de filas para cada sesión:
;;;
;;;   1. Fila RESUMEN (una por sesión):
;;;      Fecha, Archivo, Modo, Filtro, Total Entidades, Longitud Total, Unidad
;;;
;;;   2. Filas DETALLE POR CAPA (una por cada capa presente en la sesión):
;;;      Fecha, Archivo, Modo, Filtro, Capa, Entidades en Capa, Longitud en Capa, Unidad
;;;
;;; Esto permite en Excel filtrar por sesión, por archivo o por capa individualmente.
;;; ============================================================================
(defun c:metracadexport (/ path f abort)

  ; Verifica que exista al menos una medición antes de exportar
  (setq abort nil)
  (if (null *MC-HISTORY*)
    (progn
      (princ "\nNo hay datos para exportar. Ejecute METRACAD primero.")
      (setq abort T)))

  (if (not abort)
    (progn
      ; Diálogo para elegir destino del archivo CSV
      (setq path (getfiled "Exportar resultado CSV"
                            (strcat (getvar "DWGPREFIX") "metracad_export")
                            "csv" 1))
      (if path
        (progn
          (setq f (open path "w"))

          ;; --- CABECERA SECCIÓN RESUMEN ---
          (write-line "" f)
          (write-line "RESUMEN POR SESION" f)
          (write-line
            (MC:csv-row '("Fecha" "Archivo" "Modo" "Filtro"
                          "Total Entidades" "Longitud Total" "Unidad"))
            f)

          ; Una fila de resumen por cada sesión registrada
          (foreach e *MC-HISTORY*
            (write-line
              (MC:csv-row (list
                (nth 0 e)   ; Fecha
                (nth 1 e)   ; Archivo
                (nth 2 e)   ; Modo
                (nth 3 e)   ; Filtro (capa, color, o manual)
                (nth 4 e)   ; Total entidades
                (nth 5 e)   ; Longitud total
                (nth 6 e))) ; Unidad
              f))

          ;; --- CABECERA SECCIÓN DETALLE POR CAPA ---
          (write-line "" f)
          (write-line "DETALLE POR CAPA" f)
          (write-line
            (MC:csv-row '("Fecha" "Archivo" "Modo" "Filtro"
                          "Capa" "Entidades en Capa" "Longitud en Capa" "Unidad"))
            f)

          ; Para cada sesión, una fila por cada capa que contenga entidades medidas
          (foreach e *MC-HISTORY*
            (foreach ld (nth 7 e)  ; nth 7 = layer-data de esa sesión
              (write-line
                (MC:csv-row (list
                  (nth 0 e)           ; Fecha de la sesión
                  (nth 1 e)           ; Archivo DWG
                  (nth 2 e)           ; Modo
                  (nth 3 e)           ; Filtro
                  (car ld)            ; Nombre de la capa
                  (itoa (caddr ld))   ; Cantidad de entidades en esta capa
                  (MC:fmt (cadr ld))  ; Longitud acumulada en esta capa
                  (nth 6 e)))         ; Unidad
                f)))

          (close f)
          (princ (strcat "\nExportado correctamente: " path))))))

  (princ))

;;; ============================================================================
;;; COMANDO: METRACADCLIP
;;; Copia el último reporte generado al portapapeles de Windows.
;;; Útil para pegar directamente en Excel, Word, correo o cualquier editor.
;;; ============================================================================
(defun c:metracadclip (/ abort)
  (setq abort nil)

  ; Verifica que exista un reporte previo antes de intentar copiar
  (if (null *MC-LAST*)
    (progn
      (princ "\nNo hay reporte disponible. Ejecute METRACAD primero.")
      (setq abort T)))

  (if (not abort)
    (if (MC:clipboard *MC-LAST*)
      (princ "\nÚltimo reporte copiado al portapapeles.")
      (princ "\nError: no se pudo acceder al portapapeles.")))

  (princ))

;;; ============================================================================
;;; MENSAJE DE CARGA
;;; ============================================================================
(princ (strcat "\nMetraCAD v" *MC-VERSION*
               " cargado. Comandos: METRACAD  |  METRACADEXPORT  |  METRACADCLIP"))
(princ)