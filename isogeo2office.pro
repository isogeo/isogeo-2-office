FORMS = ./modules/ui/auth/ui_authentication.ui \
    	./modules/ui/credits/ui_credits.ui \
    	./modules/ui/main/ui_win_IsogeoToOffice.ui \

SOURCES = ./IsogeoToOffice.py \
		  ./modules/threads.py \
		  ./modules/export/formatter.py \
		  ./modules/export/isogeo2docx.py \
		  ./modules/export/isogeo_stats.py \
		  ./modules/utils/api.py \
		  ./modules/utils/utils.py \
		  ./modules/ui/auth/ui_authentication.py \
		  ./modules/ui/credits/ui_credits.py \
		  ./modules/ui/main/ui_win_IsogeoToOffice.py \
		  ./modules/ui/systray/ui_systraymenu.py \

TRANSLATIONS = ./i18n/IsogeoToOffice_fr.ts \
	           ./i18n/IsogeoToOffice_en.ts
