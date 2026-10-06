```
Le classeur Excel contient des onglets pour tous les secteurs. Je veux que tu génère un script Python permettant d'exclure des applications dans tous les onglets des secteurs en se basant sur des règles précises que je défini plus bas ici. Le script doit créer (une seule fois) un nouvel onglet libellé "Excluded Regle01" pour y déposer les lignes exclues. Je ne veux pas perdre de données. Si on supprime une ligne dans un onglet, celle-ci doit obligatoirement se retrouver dans "Excluded Regle01". Déplacer la ligne du secteur vers "Excluded Regle01" sans aucune transformation.

Voici la première règle :
Règle 1 — Exclure tout ce qui est installé par défaut avec Windows 11 incluant les mises à niveau  
Détail de la règle — Composantes installées par défaut avec Windows 11 et mises à niveau

Correspondances insensibles à la casse, appliquées sur la colonne « Logiciel » (et sur « Manufacturier » pour les motifs de composantes système).

A) Mises à jour et mises à niveau Windows :
- Motif « KB » suivi d'au moins 6 chiffres (ex. KB5058526, KB5124042);
- « Security Update for Microsoft », « Cumulative Update », « Hotfix », « Servicing Stack Update », « Definition Update », « Update for Microsoft »;
- Outils de mise à niveau Windows : « Assistant d'installation de Windows 11 », « Assistant Mise à jour de Windows 10 », « Microsoft Update Health Tools ».

B) Runtimes et composantes système Windows :
- « WindowsAppRuntime.* », « WinAppRuntime.* », « Windows App Runtime »;
- « Microsoft UI Xaml » (ex. Microsoft UI Xaml 2.x, Microsoft UI Xaml CBS);
- « Widgets Platform Runtime »;
- « Windows Feature Experience Pack » / « Pack d'expérience de fonctionnalités Windows »;
- « Windows Web Experience Pack »;
- « WindowsWorkload » / « Windows Workload »;
- « Windows SDK » / « Windows Software Development Kit »;
- « OEM » (mot complet);
- Composantes système Microsoft : « Microsoft Edge WebView », « Microsoft WebView », « Microsoft Visual C++ », « Microsoft Client Web », « Microsoft CrossDevice » / « Microsoft Cross Device », « Microsoft.Windows ».

C) Applications fournies avec Windows 11 (in-box), incluant les variantes localisées en français (correspondance par nom) :
- In-box « Microsoft.* » : Microsoft Paint / MSPaint, Microsoft Store, Microsoft StorePurchaseApp, Microsoft People, Microsoft Messaging, Microsoft Pay, Microsoft OneConnect, Microsoft Solitaire Collection, Microsoft Family / Family Safety, Microsoft GetHelp, Microsoft Getstarted, Microsoft Content, Microsoft ScreenSketch, Microsoft StickyNotes, Microsoft Journal, Microsoft PCManager, Microsoft 3D Viewer, Microsoft MixedReality (Portal/Link), Microsoft Todos, Microsoft StartExperiencesApp;
- Jeux et multimédia du système : Xbox (nom exact), Xbox Game UI, Xbox TCUI, Xbox Game Bar Plugin, Game Bar, Xbox One SmartGlass, Solitaire & Casual Games;
- Système : Cortana, Narrator / Narrateur (nom exact), Internet Explorer, Windows Sandbox, Windows Subsystem for Linux (+ Update, WSLg Preview), Windows Hello Setup / Configuration de Windows Hello, Windows Search, Windows Advanced Settings, Windows Default Lock Screen / Écran de verrouillage par défaut (de) Windows;
- Entrées système localisées (noms exacts) : Explorateur de fichiers / File Explorer, Paramètres / Settings, E-mail et comptes / Email and accounts / Messagerie et comptes, Compte professionnel ou scolaire / Work or school account, Galerie de photos, Photos (Preview), Impression Windows / Windows Print / File d'impression / Print Queue, Expérience Windows Shell, Contrôle d'intégrité du PC Windows, Contrôle parental Microsoft / Fonctionnalités de la famille Microsoft, Application de verrouillage d'accès attribué / pour accès affecté / Assigned Access Lock app, Eye Control / Contrôle visuel, Supprimer l'appareil en toute sécurité / Safely Remove Device, Captive Portal Flow / Débit du portail captif, Network Connection Flow / Débit de la connexion réseau, Examen / Passer un examen, Appeler, Aperçu de code-barres Windows / Windows Barcode Preview, Credential Dialog / Boîte de dialogue d'informations d'identification, Programme de résolution de l'application / des applis, PinningConfirmationDialog, Add Folder Suggestions dialog / Boîte de dialogue Ajouter des suggestions de dossier;
- Composantes techniques in-box : InputApp, AsyncTextService, App Resolver, WebPrintingService, NcsiUwpApp, LocalServiceComponents, OnnxRuntime / PSOnnxRuntime, PSTokenizer / PSTokenizerShared, WinMLShared, Engagement Framework, AIFabric (CBS);
- Packs de langue et d'expérience locale : Local Experience Pack / Module d'expérience locale, Speech Pack, Speech Translation Pack, Ink.Handwriting.

D) Outils, pilotes, agents et composantes diverses (ajouts) — motifs sans numéro de version afin d'exclure toutes les versions lorsque la version fait partie du libellé :
- Outils éditeurs, agents et services : Acrobat Notification Client, Adobe Refresh Manager, Antidote - Connectix <version>, Antidote <version> - Module français, Logi Download Assistant, Logiciel Logitech Unifying <version>, Mozilla Maintenance Service, Autodesk CER, GLPI Agent <version>, ScreenConnect Client (<hash>), mstsc-<GUID>, Assistant de téléchargement, Bonjour (nom exact), UDK Package (nom exact);
- Pilotes et utilitaires matériels : ASUS OSD Driver, ASUSPCAssistant (ex. B9ECED6F.ASUSPCAssistant), DisplayLink Graphics, doseBadge Reader Driver, DTS Audio Processing, Elan PointStick Settings, Intel Graphics Experience (ex. AppUp.IntelGraphicsExperience), Intel(R) Management and Security Status, Realtek Audio Control;
- Runtimes .NET / ASP.NET (toutes versions) : Microsoft .NET Host et Host FX Resolver, Microsoft .NET Runtime, Microsoft .Net Native Framework/Runtime Package, Microsoft ASP.NET Core Shared Framework, Microsoft Windows Desktop Runtime;
- Extensions média Microsoft : AV1VideoExtension, AVCEncoderVideoExtension, HEIFImageExtension, HEVCVideoExtension(FirstParty), MPEG2VideoExtension, RawImageExtension, VP9VideoExtensions, WebMediaExtensions, WebpImageExtension;
- Composantes et applis système Microsoft / Windows : ApplicationCompatibilityEnhancements, BingSearch / Search in Bing, CompanyPortal, D3DMappingLayers, DesktopAppInstaller, Edge GameAssist / Edge Update / Edge DevTools, Intune Management Extension, Office ActionsServer, SecHealthUI, Teams (personal) (nom exact) / Teams Meeting Add-in, Windows 11 Entreprise, AugLoop CBS, CallingShellApp, Windows Photos, StartMenuExperienceHost, WindowsAlarms, WindowsCalculator, WindowsCamera, WindowsCommunicationsApps, WindowsFeedbackHub, WindowsMaps, WindowsNotepad, WindowsTerminal, QuickAssist, MicrosoftWindows.* (Client.AIX, Client.CoreAI, CrossDevice, LKG.TwinSxS, Taskbar, Voiess, Speion, InpApp, Livtop, ...), Contenu Microsoft, Office Click-to-Run Extensibility/Localization Component, OfficePushNotificationsUtility, OneNote Virtual Printer, Local AI Manager for Microsoft 365, Windows Defender SmartScreen, VIA RoomDevice, CapturePicker, DefaultPackMSI, Flux du portail captif / Votre compte / Visionneuse web de l'application de bureau / Remote Help (noms exacts), GUID seul (ex. 1527c705-839a-4832-9118-54d4Bd6a0c89).

Les lignes exclues doivent être déplacées dans l'onglet « Excluded Regle01 » avec la raison d'exclusion suivante : Règle 1 — Installé par défaut avec Windows 11 / mise à niveau Windows.

Important — Méthode de traitement : cette règle porte uniquement sur les composantes installées par défaut avec Windows 11 et les mises à niveau (sections A à C ci-dessus). Une liste distincte des applications correspondant à cette règle doit d'abord être générée, sans déplacement de lignes. Ensuite, chaque application de cette liste est traitée séparément, par secteur. L'onglet « Excluded Regle01 » est créé une seule fois; toute ligne retirée d'un onglet de secteur doit s'y retrouver sans transformation (aucune perte de données). Le classeur est sauvegardé après chaque combinaison application/secteur pour libérer les ressources.
```
