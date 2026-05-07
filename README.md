# 📊 Excel JDBC Driver

[![Release](https://img.shields.io/badge/version-1.0.0-blue)](https://github.com/hatembenjaballah/excel-jdbc-driver/releases)

Un driver JDBC en Java qui transforme un fichier Excel en une base de données relationnelle.  
Chaque feuille devient une table, et toutes les opérations SQL courantes sont disponibles : `SELECT`, `INSERT`, `UPDATE`, `DELETE`, `CREATE TABLE`, `DROP TABLE`, jointures, agrégations, etc.

---

## ✨ Fonctionnalités

- **SQL complet** – toutes les instructions de manipulation et de définition de données.
- **Jointures** – `INNER JOIN`, `LEFT JOIN`, `RIGHT JOIN`, `FULL JOIN`.
- **Agrégations** – `COUNT`, `SUM`, `AVG`, `MIN`, `MAX` avec `GROUP BY` et `HAVING`.
- **Tri** – `ORDER BY` ascendant/descendant.
- **Filtres avancés** – `AND`, `OR`, `IN`, `IS NULL`, opérateurs de comparaison, expressions arithmétiques.
- **Types JDBC** – `VARCHAR`, `INTEGER`, `DOUBLE`, `BOOLEAN`, `DATE`, `TIMESTAMP`.
- **Métadonnées** – `DatabaseMetaData` et `ResultSetMetaData` pour l’exploration des tables.
- **Persistance automatique** – les modifications sont enregistrées dans le fichier Excel à la fermeture de la connexion.
- **Compatible DBeaver / SQuirreL SQL** – utilisation possible dans les outils graphiques.

---

## 📋 Prérequis

- **Java 8** ou version supérieure
- **Apache Maven** 3.6+
- Un fichier Excel `.xlsx` ou `.xls`

---

## 🚀 Compilation et installation

```bash
# Cloner le dépôt
git clone https://github.com/hatembenjaballah/excel-jdbc-driver.git
cd excel-jdbc-driver

# Compilation et empaquetage (uber-JAR avec dépendances)
mvn clean package