Valeurs possibles du template
=============================

| Variable remplacée dans le template | Champ Isogeo correspondant |
| :---------------------------------- | :------------------------- |
| {{ varTitle }} | (Titre de la fiche de métadonnées)[http://help.isogeo.com/fr/features/documentation/md_identification.html#titre] |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |




{{ varOwner }}
Groupe de travail




{{ varKeywordsCount }}
Nombre de mots-clés




{{ varKeywords }}
Liste des mots-clés séparés par des « ; »




{{ varAbstract }}
Résumé




{{ varDataDtCrea }}
Date de création de la donnée




{{ varDataDtUpda }}
Date de dernière modification de la donnée




{{ varDataDtPubl }}
Date de publication de la donnée




{{ varValidityStart }}
Date de début de validité




{{ varValidityEnd }}
Date de fin de validité




{{validityComment }}
Commentaire sur la période de validité




{{ varCollectContext }}
Contexte de collecte




{{ varCollectMethod }}
Méthode de collecter




{{ varNameTech }}
Nom technique (schéma.table ou fichier)




{{ varPath }}
Chemin absolu ou nom de la base de données




{{ varFormat }}
Format et version




{{ varType }}
Type de ressource




{{ varGeometry }}
Type de géométrie




{{ varObjectsCount }}
Nombre d’objets géométriques




{{ varSRS }}
Système de coordonnées (nom + EPSG)




{{ varScale }}
Échelle de référence




{{ varResolution }}
Résolution spatiale




{{ varTopologyInfo }}
Informations sur la topologie




{{ varInspireTheme }}
Thèmes INSPIRE affectés séparés par un « ; »




{{ varInspireConformity }}
Si la fiche est conforme INSPIRE ou pas




{{ varInspireLimitation }}
Limitations à la diffusion et/ou usage




{{ varContactsCount }}
Nombre de contacts affectés




{{ varContactsDetails }}
Détails des contacts




{{ varFieldsCount }}
Nombre d’attributs




{{ varMdDtCrea }}
Date de création de la métadonnée




{{ varMdDtUpda }}
Date de dernière modification de la métadonnée




{{ varMdDtExp }}
Date d’export en fichier Word





Pour les attributs, il s’agit d’un tableau qui prend 4 valeurs pour chaque attribut : name, alias, dataType, description. Exemple :
Nom (alias)
Type
Description
{%tr for i in items %}
{{ i.name }} ({{ i.alias }})
{{ i.dataType }}
{{ i.description }}
{%tr endfor %}

