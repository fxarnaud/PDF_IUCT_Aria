- telecharger les assemblies nécessaires

- Pour la partie AriaAccess 
API AriaAccess varian differente que les dll de l'esapi. Elle utilise des webservices qui sont déjà sur le serveur varian. 
Ceci est vrai pour la v15, pour la v18 ils ont créé une nouvelle api qui permet d'aller plus loin. Elle marche bien sur notre tbox v18 mais je n'ai pas encore géré le truc d'insérer des docs avec.
Je ne te promets donc pas que cette api AriaAccess avec ses webservices soit toujours sur le serveur varian.
Si tu as des erreurs qui te reviennent et qui parlent de ces webservices dis le, on pourra vérifier 2-3 trucs. Sinon c'est à voir avec l'ingé varian. Notre ingé en local (Gilles Calvet) est bien au courant (ma faute), ils peuvent communiquer si ce n'est pas le même ingé.
	-  Aller sur myvarian 
	- "myaccount" -> "API Key Management" -> "New API request"
	- Remplir toutes les infos et demander une "Software System" -> "ARIA
					et "API type "Aria Oncology services"
	-Au bout de quelques secondes on recoit en téléchargement un fichier .txt	

- Sur un pc eclipse ouvrir "Varian Service portal"
	- Relever le nom du serveur et le numéro de port grace à l'url.
	Par exemple mon url est https://srvaria15-web:55051/Portal/Account/Login?ReturnUrl=%2fPortal
	J'ai donc nom du serveur = srvaria15-web
	Numéro de port 55051
	On va se servir de ces infos après dans le code (CustomInsertDocumentsParameter.cs).
	- Aller dans "Sécurité" -> Clés API
	- Normalement tu as déjà une clé (INTEROP /AC si jene me trompe pas)
	- Cliquer sur installer et aller chercher le fichier .txt que tu as téléchargé.
	- Normalement une ligne s'ajoute avec Permet l'accès à "Infrastructure, ...., Documents
	L'install de la clé API est finie mais garde quelque part ce fichier texte on va encore s'en servir après.

- Avec ces infos dans le code, ouvrir la classe "CustomInsertDocumentsParameters"
	hostname = nom du serveur
	port = numéro de port
	dockey = ouvrir le fichier texte de la clé fournie par varian et copier-coller la partie qui suit "value="[...]

- Le chemin absolu initial est renseigné dans PDF_IUCT.cs. Modifie le pour mettre un dossier à toi.
- Toujours pour le chemin, dans le code dans DocumentGenerator.cs il utilise des images .png (il me semble) de décalages. Il faut donc que dans ton dossier tu aies un sous dossier 'images".
	Dans ce dossier on va mettre les images genre le logo de ton centre et les fichier tampons qu'il va créér et qu'on va pas voir à l'usage.
	Dans ce dossier créé un sous dossier "images". Tu peux mettre l'image de ton logo si tu veux. Il gère ça ligne 146 de "DocumentGenerator.cs".
	Idem dans ce dossier j'ai fait les images des décalages de la table qu'on a comme dans les impressions eclipse. Il gère ca ligne à partir de la ligne 683 de "DocumentGenerator.cs". 
	J'aurais pu mettre les images en statics mais à l'époque je ne savais pas le faire.

	A noter : 
		- j'ai fait en sorte de lire un fichier json (de souvenir) qui contient des mots clés pour ne pas afficher les dvh des structures qui contiennent ces mots clés là (genre les structures opt chez nous).
		Si tu as des erreurs à un moment sur ça c'est juste qu'il faut lui donner à manger un fichier. Je te file le fichier type en pj du mail, c'est bidon au besoin.
		Ca se passe dans MainViewModel.cs vers la ligne 37.
		- Il y a un système de double copie du pdf généré pour en faire un backup pour les garder pour la cyber sécurité si jamais on se fait attaquer.
		Tu peux décommenter sous le commentaire : //******Envoi sous PC tiers de sauvegarde  toujours dans MainView.xaml.cs



