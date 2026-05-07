package com.exceljdbc;

import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.List;

/**
 * Implémentation de {@link ResultSetMetaData} pour le driver Excel.
 * <p>
 * Chaque instance représente les métadonnées d'un {@link ExcelResultSet} :
 * nombre de colonnes, leurs noms, leurs libellés, et des informations
 * simplifiées sur les types SQL. Comme le driver ne type pas finement
 * les données, toutes les colonnes sont déclarées en {@link Types#VARCHAR}
 * (sauf amélioration future par inspection des données).
 * </p>
 */
public class ExcelResultSetMetaData implements ResultSetMetaData {

    /** Noms des colonnes, dans l'ordre de la requête SELECT. */
    private final List<String> columns;

    /**
     * Construit un objet de métadonnées à partir de la liste des noms de colonnes.
     *
     * @param columns les noms des colonnes (peut être vide, mais jamais {@code null})
     */
    public ExcelResultSetMetaData(List<String> columns) {
        this.columns = columns;
    }

    /**
     * Retourne le nombre total de colonnes.
     */
    @Override
    public int getColumnCount() {
        return columns.size();
    }

    /**
     * Retourne le nom de la colonne (nom original tel que dans la feuille Excel,
     * ou alias SQL). Index 1‑based.
     */
    @Override
    public String getColumnName(int column) throws SQLException {
        checkIndex(column);
        return columns.get(column - 1);
    }

    /**
     * Retourne le libellé de la colonne. Ici identique au nom.
     * @see #getColumnName(int)
     */
    @Override
    public String getColumnLabel(int column) throws SQLException {
        return getColumnName(column);
    }

    /**
     * Type SQL générique : toujours {@link Types#VARCHAR} dans cette version simplifiée.
     * Amélioration possible : analyser les valeurs pour déduire le type réel.
     */
    @Override
    public int getColumnType(int column) throws SQLException {
        checkIndex(column);
        return Types.VARCHAR;
    }

    /**
     * Nom du type SQL en clair.
     */
    @Override
    public String getColumnTypeName(int column) throws SQLException {
        return "VARCHAR";
    }

    /**
     * Nom de la classe Java correspondant au type de la colonne.
     */
    @Override
    public String getColumnClassName(int column) throws SQLException {
        return "java.lang.String";
    }

    /**
     * Taille d'affichage de la colonne (retourne 255 par défaut).
     */
    @Override
    public int getColumnDisplaySize(int column) throws SQLException {
        checkIndex(column);
        return 255;
    }

    /**
     * Précision (nombre total de chiffres pour les types numériques). Non pertinent pour VARCHAR.
     */
    @Override
    public int getPrecision(int column) throws SQLException {
        return 0;
    }

    /**
     * Échelle (nombre de chiffres après la virgule). Non pertinent ici.
     */
    @Override
    public int getScale(int column) throws SQLException {
        return 0;
    }

    /**
     * Indique si la colonne peut contenir NULL.
     * @return {@link ResultSetMetaData#columnNullable} car toutes les cellules peuvent être vides.
     */
    @Override
    public int isNullable(int column) throws SQLException {
        checkIndex(column);
        return columnNullable;
    }

    /**
     * Aucune table n'est associée à un ResultSet global (jointures, agrégats).
     * Retourne une chaîne vide.
     */
    @Override
    public String getTableName(int column) throws SQLException {
        return "";
    }

    /**
     * Pas de schéma supporté.
     */
    @Override
    public String getSchemaName(int column) throws SQLException {
        return "";
    }

    /**
     * Pas de catalogue supporté.
     */
    @Override
    public String getCatalogName(int column) throws SQLException {
        return "";
    }

    /**
     * Les colonnes ne sont pas auto-incrémentées.
     */
    @Override
    public boolean isAutoIncrement(int column) throws SQLException {
        return false;
    }

    /**
     * Les noms de colonnes sont considérés insensibles à la casse pour l'accès utilisateur.
     * (Bien que Excel conserve la casse, la recherche avec {@code findColumn} ignore la casse.)
     */
    @Override
    public boolean isCaseSensitive(int column) throws SQLException {
        return false;
    }

    /**
     * Toutes les colonnes sont utilisables dans une clause WHERE.
     */
    @Override
    public boolean isSearchable(int column) throws SQLException {
        return true;
    }

    /**
     * Pas de notion de devise.
     */
    @Override
    public boolean isCurrency(int column) throws SQLException {
        return false;
    }

    /**
     * Retourne {@code false} car les colonnes ne sont pas signées (type VARCHAR).
     */
    @Override
    public boolean isSigned(int column) throws SQLException {
        return false;
    }

    /**
     * Indique si la colonne est en lecture seule.
     * Ici on retourne {@code false} car le driver permet les mises à jour,
     * même si ce ResultSet précis est non modifiable.
     */
    @Override
    public boolean isReadOnly(int column) throws SQLException {
        return false;
    }

    /**
     * Indique si la colonne est accessible en écriture (oui).
     */
    @Override
    public boolean isWritable(int column) throws SQLException {
        return true;
    }

    /**
     * Indique si la colonne est définitivement accessible en écriture (oui).
     */
    @Override
    public boolean isDefinitelyWritable(int column) throws SQLException {
        return true;
    }

    // ─── Implémentations de wrapper ───
    @Override
    public <T> T unwrap(Class<T> iface) throws SQLException {
        throw new SQLException("Non wrapper");
    }

    @Override
    public boolean isWrapperFor(Class<?> iface) throws SQLException {
        return false;
    }

    /**
     * Vérifie que l'index de colonne est correct (entre 1 et le nombre de colonnes).
     */
    private void checkIndex(int column) throws SQLException {
        if (column < 1 || column > columns.size()) {
            throw new SQLException("Index de colonne invalide : " + column
                    + " (doit être entre 1 et " + columns.size() + ")");
        }
    }
}