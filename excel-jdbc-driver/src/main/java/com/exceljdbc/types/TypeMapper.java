package com.exceljdbc.types;

import java.sql.Types;
import java.util.Date;

/**
 * Utilitaire de mapping entre les types Java (rencontrés dans les cellules Excel)
 * et les types SQL de l'API JDBC.
 * <p>
 * Cette classe permet de déterminer le {@link java.sql.Types} correspondant
 * à une valeur Java, et vice‑versa. Elle est utilisée notamment par
 * {@code ExcelResultSetMetaData} pour renseigner le type des colonnes
 * (à la place du simple {@code VARCHAR} utilisé par défaut).
 * </p>
 */
public class TypeMapper {

    /**
     * Retourne la constante {@link java.sql.Types} la plus appropriée pour la
     * valeur Java donnée.
     *
     * @param value la valeur issue d'une cellule Excel (peut être {@code null})
     * @return le type JDBC (par défaut {@link Types#VARCHAR})
     */
    public static int getSqlType(Object value) {
        if (value == null) {
            return Types.VARCHAR; // NULL est compatible avec tous les types, on choisit VARCHAR
        }
        if (value instanceof String) {
            return Types.VARCHAR;
        }
        if (value instanceof Long || value instanceof Integer || value instanceof Short || value instanceof Byte) {
            return Types.INTEGER;
        }
        if (value instanceof Double || value instanceof Float) {
            return Types.DOUBLE;
        }
        if (value instanceof java.sql.Date) {
            return Types.DATE;
        }
        if (value instanceof java.sql.Timestamp) {
            return Types.TIMESTAMP;
        }
        if (value instanceof Date) {
            // java.util.Date sans précision → TIMESTAMP
            return Types.TIMESTAMP;
        }
        if (value instanceof Boolean) {
            return Types.BOOLEAN;
        }
        // Pour les autres (BigDecimal, etc.), on retourne VARCHAR par défaut
        return Types.VARCHAR;
    }

    /**
     * Retourne le nom lisible du type SQL correspondant à une valeur Java.
     * Utile pour {@link java.sql.ResultSetMetaData#getColumnTypeName(int)}.
     *
     * @param value la valeur (ou {@code null}) pour laquelle on veut le nom du type
     * @return le nom du type SQL (ex: "INTEGER", "VARCHAR", "DOUBLE"...)
     */
    public static String getSqlTypeName(Object value) {
        int type = getSqlType(value);
        return getTypeName(type);
    }

    /**
     * Convertit une constante {@link java.sql.Types} en son nom lisible.
     *
     * @param sqlType constante JDBC (ex: {@link Types#INTEGER})
     * @return le nom du type (ex: "INTEGER")
     */
    public static String getTypeName(int sqlType) {
        switch (sqlType) {
            case Types.VARCHAR:    return "VARCHAR";
            case Types.INTEGER:    return "INTEGER";
            case Types.DOUBLE:     return "DOUBLE";
            case Types.BOOLEAN:    return "BOOLEAN";
            case Types.DATE:       return "DATE";
            case Types.TIMESTAMP:  return "TIMESTAMP";
            default:               return "VARCHAR";
        }
    }

    /**
     * Convertit un nom de type SQL (tel qu'utilisé dans un CREATE TABLE) en
     * constante {@link java.sql.Types}.
     *
     * @param typeName le nom du type (insensible à la casse, ex: "int", "VARCHAR")
     * @return la constante JDBC correspondante, ou {@link Types#VARCHAR} par défaut
     */
    public static int getSqlTypeFromName(String typeName) {
        if (typeName == null) return Types.VARCHAR;
        String upper = typeName.toUpperCase().trim();
        switch (upper) {
            case "INT": case "INTEGER":   return Types.INTEGER;
            case "BIGINT": case "LONG":   return Types.INTEGER; // approximé
            case "DOUBLE": case "FLOAT":
            case "REAL": case "NUMERIC":
            case "DECIMAL":               return Types.DOUBLE;
            case "VARCHAR": case "CHAR":
            case "TEXT": case "STRING":   return Types.VARCHAR;
            case "BOOLEAN": case "BOOL":  return Types.BOOLEAN;
            case "DATE":                  return Types.DATE;
            case "TIMESTAMP":
            case "DATETIME":              return Types.TIMESTAMP;
            default:                      return Types.VARCHAR;
        }
    }

    /**
     * Tente de convertir une valeur Java brute en un type adapté à l'écriture
     * dans une cellule Excel. Par exemple, un {@link java.util.Date} sera converti
     * en {@link java.sql.Timestamp} pour conserver l'heure.
     *
     * @param value la valeur à normaliser
     * @return la valeur normalisée (ou la valeur d'origine si aucun traitement)
     */
    public static Object normalizeValue(Object value) {
        if (value == null) return null;
        // Conserver les dates sous forme java.util.Date (POI les gère)
        if (value instanceof Date) {
            return value;
        }
        if (value instanceof Boolean) {
            return value;
        }
        if (value instanceof Number) {
            return value;
        }
        if (value instanceof String) {
            return value;
        }
        // Pour tout autre type, on retourne la représentation string
        return value.toString();
    }
}