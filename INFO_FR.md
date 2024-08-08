La gestion des droits d'accès et des comptes utilisateurs devient de plus en plus facile avec le connecteur CPAGE L'intégration CPAGE HelloID permet d'automatiser ces processus en se basant sur le système RH. Dans cet article, je donne plus d'explications sur ce qu'implique cette intégration et les possibilités spécifiques qu'elle offre.

## Qu'est-ce que CPAGE ?
Le logiciel CPAGE RH est un outil de gestion des ressources humaines spécifiquement conçu pour le secteur public en France, en particulier pour les établissements de santé et les administrations.

## Pourquoi l'intégration de CPAGE est-elle utile ?

La gestion des comptes d'utilisateurs et des droits d'accès peut prendre du temps et être sujette à des erreurs, en particulier lorsqu'il s'agit de processus manuels. Il est essentiel de s'assurer que les nouveaux employés disposent des autorisations correctes dès le premier jour afin qu'ils puissent commencer à travailler immédiatement. De plus, il est important que les changements de rôles, de lieux ou d'unité founctionnelle (UF) soient reflétés avec précision dans le système. L'intégration de CPAGE HelloID permet d'automatiser ces processus sur la base du système RH. En outre, le connecteur CPAGE offre une intégration avec des systèmes cibles courants tels que :

* Active Directory
* Entra ID (Azure AD)
* GLPI
* Easily, Hopital Manger, Cariatides
* Zimbra, Google, Exchange
* Chronos

De plus amples détails sur l'intégration avec ces systèmes cibles sont disponibles plus loin dans l'article.

## L'intégration de HelloID CPAGE vous permet :
👉 **La création accélérée de comptes :** Les processus automatisés accélèrent le temps nécessaire à la création de nouveaux comptes. Les nouveaux employés peuvent être productifs dès le premier jour, avec une intervention manuelle minimale.

👉 **Une gestion des comptes sans erreur :** Les processus automatisés réduisent le risque d'erreurs dans la gestion des comptes. Les employés disposent toujours exactement des comptes et des droits auxquels ils ont droit selon la matrice d'autorisation.

👉 **Une synchronisation bidirectionnelle :** Les modifications apportées dans CPAGE sont automatiquement détectées et traitées dans tous les systèmes et applications liés à HelloID. Les noms d'utilisateur et les adresses électroniques générés par HelloID peuvent également être automatiquement retranscrits dans CPAGE.

👉 **Une amélioration des niveaux de service et de la sécurité :** L'intégration de CPAGE avec HelloID aide les organisations à respecter les accords de niveau de service (SLA) et à passer les audits informatiques.

## Comment HelloID s'intègre à CPAGE
CPAGE peut être intégré avec HelloID en tant que connecteur source et cible. Grâce à cette intégration, HelloID détecte tout changement dans CPAGE et gère automatiquement les comptes utilisateurs dans le paysage applicatif conformément aux règles et procédures définies.

HelloID prend en charge divers processus du cycle de vie de l'identité, notamment la création, la modification et la suppression de comptes d'utilisateurs sur la base des informations contenues dans CPAGE. Le tableau ci-dessous donne quelques exemples de changements courants qui peuvent se produire dans CPAGE et les procédures associées.

| Évènement                                | Procédure dans les systèmes cibles |
| ---------------------------------------- | ---------------------------------- |
| **Nouvel employé**                       | Sur la base des informations contenues dans CPAGE, un compte d'utilisateur est créé dans les applications liées, avec les appartenances de groupe correctes. En fonction du rôle du nouvel employé, des comptes d'utilisateur sont créés et des droits attribués dans d'autres systèmes. |
| **Changement de fonction**               | Les comptes d'utilisateurs reçoivent des droits différents dans les systèmes liés. Le modèle d'autorisation de HelloID est automatiquement consulté, et les droits sont ajoutés et supprimés. |
| **Changement d'UF**                      | Le compte d'utilisateur est déplacé vers une autre OU dans AD et doté de droits spécifiques à son unité fonctionelle. |
| **Départ de l'établissement**            | Les comptes d'utilisateurs sont désactivés et les employés concernés en sont informés. Après un certain temps, les comptes sont automatiquement supprimés. |

Le connecteur CPAGE permet d'échanger de manière transparente des données entre CPAGE et HelloID, telles que des données sur le personnel, des informations sur les contrats et des données organisationnelles. Le système d'application-connecteur de CPAGE permet la connexion entre les deux systèmes à l'aide de requêtes en transaction SQL.

## Lier CPAGE à Active Directory, Azure, GLPI, Easily, etc.
HelloID permet de relier facilement CPAGE à des systèmes cibles communs. Cela permet d'améliorer l'automatisation et la collaboration entre les différentes applications. Voici quelques exemples d'intégrations:
* **CPAGE - Intégration d'Active Directory**
Avec HelloID, vous bénéficiez d'une synchronisation transparente entre CPAGE et Active Directory (AD), ce qui permet de maintenir les comptes et les droits d'accès à jour à tout moment.
* **CPAGE - Intégration d'Azure AD**
Bénéficiez d'une synchronisation automatisée entre CPAGE et Azure AD avec HelloID. Basé sur CPAGE, HelloID assure la création automatique des comptes, la gestion des droits et la gestion des licences (par exemple, les licences Office).
* **CPAGE - Intégration GLPI**
Augmentez l'efficacité entre les RH et l'informatique en reliant CPAGE à GLPI. HelloID peut, par exemple, créer automatiquement des tickets lors de l'intégration, ce qui permet de rationaliser le processus.

En outre, HelloID prend en charge plus de 200 connecteurs, ce qui offre un large éventail de possibilités d'intégration entre CPAGE et d'autres systèmes.
