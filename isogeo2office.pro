FORMS = ./modules/ui/auth/ui_authentication.ui \
    	./modules/ui/credits/ui_credits.ui \
    	./modules/ui/main/ui_IsogeoToOffice.ui \

SOURCES = ./__main__.py \
		  ./isogeo2office.py \
		  ./modules/export/formatter.py \
		  ./modules/export/isogeo2docx.py \
		  ./modules/export/isogeo2xlsx.py \
		  ./modules/export/isogeo_stats.py \
		  ./modules/utils/utils.py \

TRANSLATIONS = ./i18n/IsogeoToOffice_fr.ts \
	           ./i18n/IsogeoToOffice_en.ts
