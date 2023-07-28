- telecharger les assemblies : 

- Aller sur myvarian 
	- Cliquer sur "myaccount" -> "API Key Management" -> "New API request"
	- Remplir toutes les infos et demander une "Software System" -> "ARIA
					et "API type "Aria Oncology services"
	-Au bout de quelques secondes on recoit en téléchargement un fichier .txt
	

- Sur un pc eclipse ouvrir "Varian Service portal"
	- Relever le nom du serveur et le numéro de port grace à l'url.
	Par exemple mon url est https://srvaria15-web:55051/Portal/Account/Login?ReturnUrl=%2fPortal
	J'ai donc nom du serveur = srvaria15-wab
	Numéro de port 55051
	On va se servir de ces infos après.
	- Aller dans "Sécurité" -> Clés API
	- Normalement tu as déjà une clé (INTEROP /AC si jene me trompe pas)
	- Cliquer sur installer et aller chercher le fichier .txt que tu as téléchargé.
	- Normalement une ligne s'ajoute avec Permet l'accès à "Infrastructure, ...., Documents
	L'install de la clé API est finie mais garde quelque part ce fichier texte on va encore s'en servir après.

- Avec ces infos dans le code, ouvrir la classe "CustomInsertDocumentsParameters"
	hostname = nom du serveur
	port = numéro de port
	dockey = ouvrir le fichier texte de la clé fournie par varian et copier-coller la partie qui suit "value="[...]

- Encore dans le code, dans "Script.cs" modifie le nom du dossier et pointe vers un dossier que tu créés.
	Dans ce dossier on va mettre les images genre le logo de ton centre et les fichier tampons qu'il va créér et qu'on va pas voir à l'usage.
	Dans ce dossier créé un sous dossier "images". Tu peux mettre l'image de ton logo si tu veux. Il gère ça ligne 146 de "DocumentGenerator.cs".
	Idem dans ce dossier j'ai fait les images des décalages de la table qu'on a comme dans les impressions eclipse. Il gère ca ligne à partir de la ligne 683 de "DocumentGenerator.cs". 



