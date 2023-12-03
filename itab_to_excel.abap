*&---------------------------------------------------------------------*
*& Report ZBC_ITAB_TO_EXCEL
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT ZBC_ITAB_TO_EXCEL.
TYPES: BEGIN OF gty_header,
         line type char20,
       END OF gty_header.


DATA: gv_filename type STRING,
      gt_scarr    type TABLE of scarr,
      gs_header   type gty_header,
      gt_header   type table of gty_header.

PARAMETERS: p_path type string. " kaydetme yolu icin parametre tan�mlad�k.

at SELECTION-SCREEN on VALUE-REQUEST FOR p_path. " search help ekledik
  CALL METHOD cl_gui_frontend_services=>DIRECTORY_BROWSE " kaydetme yolunu bu fonksiyona vercez ve bu kaydedecek.
*    EXPORTING
*      WINDOW_TITLE         =                  " Title of Browsing Window
*      INITIAL_FOLDER       =                  " Start Browsing Here
    CHANGING
      SELECTED_FOLDER      = p_path                " Folder Selected By User
    EXCEPTIONS
      CNTL_ERROR           = 1                " Control error
      ERROR_NO_GUI         = 2                " No GUI available
      NOT_SUPPORTED_BY_GUI = 3                " GUI does not support this
      OTHERS               = 4.
  IF SY-SUBRC <> 0.
*   MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*     WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.



START-OF-SELECTION.

  select * from scarr INTO TABLE gt_scarr.

  CONCATENATE p_path '\' sy-datum '-' sy-uzeit '.xls' INTO gv_filename. " secilen kaydetme yolu bu degiskene atand� ve dosya ismi belirlendi.

  gs_header-line = 'mandt'.
  append gs_header to gt_header.
  gs_header-LINE = 'K�sa Tan�m'.
  append gs_header to gt_header.
  gs_header-line = 'H� Ad'.
  append gs_header to gt_header.
  gs_header-line = 'PR'.
  APPEND gs_header to gt_header.
  gs_header-LINE = 'URL'.
  APPEND gs_header to gt_header. " kolon basl�klar�n� bi tablodan alacag� icin bu tabllonun icini s�ras�yla kolon basl�klar� olarak ayarlad�k.

  CALL FUNCTION 'GUI_DOWNLOAD'
    EXPORTING
      FILENAME              = GV_FILENAME
      FILETYPE              = 'ASC' "excel tipinde indirme yapacag�m�z icin asc tipi olmal�
      WRITE_FIELD_SEPARATOR = 'X' " buray� secmezsek eger her bir sat�r� tek bir h�creye yazacak ama buray� secersek h�cre h�cre exceli kaydedecek
    TABLES
      DATA_TAB              = gt_scarr
      FIELDNAMES            = gt_header. "buras� da kolon basl�g� verebilmemiz icin yap�lm�s
