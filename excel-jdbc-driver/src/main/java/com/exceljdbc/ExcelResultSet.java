package com.exceljdbc;

import java.io.InputStream;
import java.io.Reader;
import java.math.BigDecimal;
import java.net.URL;
import java.sql.*;
import java.util.*;
import java.util.Date;

/**
 * Implémentation de {@link ResultSet} pour le driver Excel.
 * <p>
 * Un {@code ExcelResultSet} est construit à partir d'une liste de lignes,
 * chaque ligne étant une {@link Map} associant nom de colonne → valeur.
 * Il fournit un accès en lecture seule avec un curseur progressant
 * uniquement vers l'avant ({@code TYPE_FORWARD_ONLY}). 
 * </p>
 */
public class ExcelResultSet implements ResultSet {

    /** Les données du résultat : chaque élément est une ligne. */
    private final List<Map<String, Object>> rows;

    /** Position actuelle du curseur dans {@link #rows} (−1 avant le premier appel à {@code next()}). */
    private int cursor = -1;

    /** État fermé du ResultSet. */
    private boolean closed = false;

    /** Métadonnées (noms et types des colonnes). */
    private final ResultSetMetaData metaData;

    /** Drapeau indiquant si la dernière valeur lue était {@code null}. */
    private boolean wasNull = false;

    /**
     * Construit un {@code ExcelResultSet} avec les données fournies.
     *
     * @param rows    liste des lignes (chaque {@code Map} associe nom de colonne → valeur)
     * @param columns noms des colonnes (dans l'ordre d'affichage)
     * @throws SQLException si les métadonnées ne peuvent être créées
     */
    public ExcelResultSet(List<Map<String, Object>> rows, List<String> columns) throws SQLException {
        this.rows = rows != null ? rows : new ArrayList<>();
        this.metaData = new ExcelResultSetMetaData(columns);
    }

    // ── Déplacement du curseur ─────────────────────────────────────────────

    /**
     * Avance le curseur d'une ligne. Retourne {@code false} s'il n'y a plus de lignes.
     */
    @Override
    public boolean next() throws SQLException {
        checkClosed();
        cursor++;
        return cursor < rows.size();
    }

    /**
     * Positionne le curseur avant la première ligne (compatible mais non utilisé en mode forward-only).
     */
    @Override
    public void beforeFirst() throws SQLException {
        checkClosed();
        cursor = -1;
    }

    /**
     * Positionne le curseur après la dernière ligne.
     */
    @Override
    public void afterLast() throws SQLException {
        checkClosed();
        cursor = rows.size();
    }

    /**
     * Positionne le curseur sur la première ligne.
     * @return {@code true} si une ligne existe, {@code false} si le ResultSet est vide.
     */
    @Override
    public boolean first() throws SQLException {
        checkClosed();
        if (rows.isEmpty()) return false;
        cursor = 0;
        return true;
    }

    /**
     * Positionne le curseur sur la dernière ligne.
     */
    @Override
    public boolean last() throws SQLException {
        checkClosed();
        if (rows.isEmpty()) return false;
        cursor = rows.size() - 1;
        return true;
    }

    /** Retourne le numéro de la ligne courante (1 based) ou 0 si positionné avant/après. */
    @Override
    public int getRow() {
        return (cursor >= 0 && cursor < rows.size()) ? cursor + 1 : 0;
    }

    /** Retourne la position absolue du curseur (non supporté → lance SQLException). */
    @Override
    public boolean absolute(int row) throws SQLException {
        throw new SQLFeatureNotSupportedException("absolute non supporté");
    }

    /** Déplacement relatif (non supporté). */
    @Override
    public boolean relative(int rows) throws SQLException {
        throw new SQLFeatureNotSupportedException("relative non supporté");
    }

    @Override
    public boolean previous() throws SQLException {
        throw new SQLFeatureNotSupportedException("previous non supporté (forward only)");
    }

    @Override
    public boolean isBeforeFirst() throws SQLException {
        checkClosed();
        return cursor == -1;
    }

    @Override
    public boolean isAfterLast() throws SQLException {
        checkClosed();
        return cursor >= rows.size();
    }

    @Override
    public boolean isFirst() throws SQLException {
        checkClosed();
        return cursor == 0 && !rows.isEmpty();
    }

    @Override
    public boolean isLast() throws SQLException {
        checkClosed();
        return cursor == rows.size() - 1 && !rows.isEmpty();
    }

    // ── Méthodes pour récupérer la valeur par index ou nom de colonne ─────

    // ----- Par index (1 based) -----
    @Override
    public String getString(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        return val != null ? val.toString() : null;
    }

    @Override
    public boolean getBoolean(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        if (val instanceof Boolean) return (Boolean) val;
        if (val instanceof String) return Boolean.parseBoolean((String) val);
        return false;
    }

    @Override
    public byte getByte(int columnIndex) throws SQLException {
        return (byte) getLong(columnIndex);
    }

    @Override
    public short getShort(int columnIndex) throws SQLException {
        return (short) getLong(columnIndex);
    }

    @Override
    public int getInt(int columnIndex) throws SQLException {
        return (int) getLong(columnIndex);
    }

    @Override
    public long getLong(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        if (val instanceof Number) return ((Number) val).longValue();
        if (val instanceof String) return Long.parseLong((String) val);
        return 0;
    }

    @Override
    public float getFloat(int columnIndex) throws SQLException {
        return (float) getDouble(columnIndex);
    }

    @Override
    public double getDouble(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        if (val instanceof Number) return ((Number) val).doubleValue();
        if (val instanceof String) return Double.parseDouble((String) val);
        return 0.0;
    }

    @Override
    public BigDecimal getBigDecimal(int columnIndex, int scale) throws SQLException {
        double d = getDouble(columnIndex);
        if (wasNull) return null;
        return BigDecimal.valueOf(d).setScale(scale, BigDecimal.ROUND_HALF_UP);
    }

    @Override
    public byte[] getBytes(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        if (val instanceof byte[]) return (byte[]) val;
        if (val instanceof String) return ((String) val).getBytes();
        return null;
    }

    @Override
    public java.sql.Date getDate(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        if (val instanceof Date) return new java.sql.Date(((Date) val).getTime());
        if (val instanceof java.sql.Date) return (java.sql.Date) val;
        // essayer de parser une chaîne (simplifié)
        if (val instanceof String) {
            try {
                return java.sql.Date.valueOf((String) val);
            } catch (IllegalArgumentException e) { /* ignore */ }
        }
        return null;
    }

    @Override
    public Time getTime(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        if (val instanceof Time) return (Time) val;
        if (val instanceof String) {
            try {
                return Time.valueOf((String) val);
            } catch (IllegalArgumentException e) { }
        }
        return null;
    }

    @Override
    public Timestamp getTimestamp(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        if (val instanceof Timestamp) return (Timestamp) val;
        if (val instanceof Date) return new Timestamp(((Date) val).getTime());
        if (val instanceof String) {
            try {
                return Timestamp.valueOf((String) val);
            } catch (IllegalArgumentException e) { }
        }
        return null;
    }

    @Override
    public Object getObject(int columnIndex) throws SQLException {
        Object val = getValue(columnIndex);
        wasNull = (val == null);
        return val;
    }

    // ----- Par nom de colonne -----
    @Override
    public String getString(String columnLabel) throws SQLException {
        return getString(findColumn(columnLabel));
    }

    // Les autres méthodes getXxx(String) délèguent simplement à la version avec index.
    @Override
    public boolean getBoolean(String columnLabel) throws SQLException {
        return getBoolean(findColumn(columnLabel));
    }
    @Override
    public byte getByte(String columnLabel) throws SQLException {
        return getByte(findColumn(columnLabel));
    }
    @Override
    public short getShort(String columnLabel) throws SQLException {
        return getShort(findColumn(columnLabel));
    }
    @Override
    public int getInt(String columnLabel) throws SQLException {
        return getInt(findColumn(columnLabel));
    }
    @Override
    public long getLong(String columnLabel) throws SQLException {
        return getLong(findColumn(columnLabel));
    }
    @Override
    public float getFloat(String columnLabel) throws SQLException {
        return getFloat(findColumn(columnLabel));
    }
    @Override
    public double getDouble(String columnLabel) throws SQLException {
        return getDouble(findColumn(columnLabel));
    }
    @Override
    public BigDecimal getBigDecimal(String columnLabel, int scale) throws SQLException {
        return getBigDecimal(findColumn(columnLabel), scale);
    }
    @Override
    public byte[] getBytes(String columnLabel) throws SQLException {
        return getBytes(findColumn(columnLabel));
    }
    @Override
    public java.sql.Date getDate(String columnLabel) throws SQLException {
        return getDate(findColumn(columnLabel));
    }
    @Override
    public Time getTime(String columnLabel) throws SQLException {
        return getTime(findColumn(columnLabel));
    }
    @Override
    public Timestamp getTimestamp(String columnLabel) throws SQLException {
        return getTimestamp(findColumn(columnLabel));
    }
    @Override
    public Object getObject(String columnLabel) throws SQLException {
        return getObject(findColumn(columnLabel));
    }

    // ── Travail sur la ligne courante ────────────────────────────────────

    /**
     * Retourne la valeur de la colonne à l'index spécifié (1‑based) dans la ligne courante.
     */
    private Object getValue(int columnIndex) throws SQLException {
        if (cursor < 0 || cursor >= rows.size()) {
            throw new SQLException("Curseur invalide : positionnez-vous sur une ligne valide");
        }
        Map<String, Object> row = rows.get(cursor);
        String colName = metaData.getColumnName(columnIndex);
        return row.get(colName);
    }

    /**
     * Recherche l'index (1‑based) d'une colonne par son nom (insensible à la casse).
     */
    @Override
    public int findColumn(String columnLabel) throws SQLException {
        for (int i = 1; i <= metaData.getColumnCount(); i++) {
            if (metaData.getColumnName(i).equalsIgnoreCase(columnLabel)) {
                return i;
            }
        }
        throw new SQLException("Colonne introuvable : " + columnLabel);
    }

    @Override
    public boolean wasNull() {
        return wasNull;
    }

    // ── Métadonnées et état ──────────────────────────────────────────────

    @Override
    public ResultSetMetaData getMetaData() {
        return metaData;
    }

    @Override
    public Statement getStatement() {
        return null; // pas de référence au Statement pour simplifier
    }

    @Override
    public SQLWarning getWarnings() {
        return null;
    }

    @Override
    public void clearWarnings() {
    }

    @Override
    public String getCursorName() throws SQLException {
        return null;
    }

    @Override
    public boolean isClosed() {
        return closed;
    }

    @Override
    public void close() {
        closed = true;
    }

    private void checkClosed() throws SQLException {
        if (closed) {
            throw new SQLException("ResultSet fermé");
        }
    }

    // ── Colonnes non supportées (lecture de flux, caractères, objets exotiques) ─
    @Override
    public InputStream getAsciiStream(int columnIndex) throws SQLException {
        throw new SQLFeatureNotSupportedException("AsciiStream non supporté");
    }
    @Override
    public InputStream getUnicodeStream(int columnIndex) throws SQLException {
        throw new SQLFeatureNotSupportedException("UnicodeStream non supporté");
    }
    @Override
    public InputStream getBinaryStream(int columnIndex) throws SQLException {
        throw new SQLFeatureNotSupportedException("BinaryStream non supporté");
    }
    @Override
    public Reader getCharacterStream(int columnIndex) throws SQLException {
        throw new SQLFeatureNotSupportedException("CharacterStream non supporté");
    }
    @Override
    public InputStream getAsciiStream(String columnLabel) throws SQLException {
        return getAsciiStream(findColumn(columnLabel));
    }
    // ... autres méthodes similaires implémentées avec des throw

    // ── Méthodes de mise à jour (ResultSet modifiable) → non supportées ──
    @Override
    public void updateNull(int columnIndex) throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    @Override
    public void updateBoolean(int columnIndex, boolean x) throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    // Tous les autres updateXxx (updateString, updateInt, etc.) sont similaires.
    @Override
    public void updateRow() throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    @Override
    public void insertRow() throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    @Override
    public void deleteRow() throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    @Override
    public void refreshRow() throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    @Override
    public void cancelRowUpdates() throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    @Override
    public void moveToInsertRow() throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }
    @Override
    public void moveToCurrentRow() throws SQLException {
        throw new SQLFeatureNotSupportedException("ResultSet non modifiable");
    }

    @Override
    public boolean rowUpdated() throws SQLException { return false; }
    @Override
    public boolean rowInserted() throws SQLException { return false; }
    @Override
    public boolean rowDeleted() throws SQLException { return false; }

    // Méthodes de wrapper
    @Override
    public <T> T unwrap(Class<T> iface) throws SQLException {
        throw new SQLException("Non wrapper");
    }

    @Override
    public boolean isWrapperFor(Class<?> iface) throws SQLException {
        return false;
    }

    // ── Diverses méthodes par défaut ──────────────────────────────────────
    @Override
    public int getType() throws SQLException {
        return ResultSet.TYPE_FORWARD_ONLY;
    }
    @Override
    public int getConcurrency() throws SQLException {
        return ResultSet.CONCUR_READ_ONLY;
    }
    @Override
    public int getHoldability() throws SQLException {
        return ResultSet.CLOSE_CURSORS_AT_COMMIT;
    }
    @Override
    public void setFetchDirection(int direction) throws SQLException { /* ignoré */ }
    @Override
    public int getFetchDirection() { return ResultSet.FETCH_FORWARD; }
    @Override
    public void setFetchSize(int rows) { }
    @Override
    public int getFetchSize() { return 0; }

	@Override
	public InputStream getUnicodeStream(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public InputStream getBinaryStream(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Reader getCharacterStream(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public BigDecimal getBigDecimal(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public BigDecimal getBigDecimal(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void updateByte(int columnIndex, byte x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateShort(int columnIndex, short x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateInt(int columnIndex, int x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateLong(int columnIndex, long x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateFloat(int columnIndex, float x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateDouble(int columnIndex, double x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBigDecimal(int columnIndex, BigDecimal x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateString(int columnIndex, String x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBytes(int columnIndex, byte[] x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateDate(int columnIndex, java.sql.Date x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateTime(int columnIndex, Time x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateTimestamp(int columnIndex, Timestamp x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateAsciiStream(int columnIndex, InputStream x, int length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBinaryStream(int columnIndex, InputStream x, int length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateCharacterStream(int columnIndex, Reader x, int length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateObject(int columnIndex, Object x, int scaleOrLength) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateObject(int columnIndex, Object x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNull(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBoolean(String columnLabel, boolean x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateByte(String columnLabel, byte x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateShort(String columnLabel, short x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateInt(String columnLabel, int x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateLong(String columnLabel, long x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateFloat(String columnLabel, float x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateDouble(String columnLabel, double x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBigDecimal(String columnLabel, BigDecimal x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateString(String columnLabel, String x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBytes(String columnLabel, byte[] x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateDate(String columnLabel, java.sql.Date x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateTime(String columnLabel, Time x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateTimestamp(String columnLabel, Timestamp x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateAsciiStream(String columnLabel, InputStream x, int length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBinaryStream(String columnLabel, InputStream x, int length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateCharacterStream(String columnLabel, Reader reader, int length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateObject(String columnLabel, Object x, int scaleOrLength) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateObject(String columnLabel, Object x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public Object getObject(int columnIndex, Map<String, Class<?>> map) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Ref getRef(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Blob getBlob(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Clob getClob(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Array getArray(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Object getObject(String columnLabel, Map<String, Class<?>> map) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Ref getRef(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Blob getBlob(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Clob getClob(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Array getArray(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public java.sql.Date getDate(int columnIndex, Calendar cal) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public java.sql.Date getDate(String columnLabel, Calendar cal) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Time getTime(int columnIndex, Calendar cal) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Time getTime(String columnLabel, Calendar cal) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Timestamp getTimestamp(int columnIndex, Calendar cal) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Timestamp getTimestamp(String columnLabel, Calendar cal) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public URL getURL(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public URL getURL(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void updateRef(int columnIndex, Ref x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateRef(String columnLabel, Ref x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBlob(int columnIndex, Blob x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBlob(String columnLabel, Blob x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateClob(int columnIndex, Clob x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateClob(String columnLabel, Clob x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateArray(int columnIndex, Array x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateArray(String columnLabel, Array x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public RowId getRowId(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public RowId getRowId(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void updateRowId(int columnIndex, RowId x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateRowId(String columnLabel, RowId x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNString(int columnIndex, String nString) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNString(String columnLabel, String nString) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNClob(int columnIndex, NClob nClob) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNClob(String columnLabel, NClob nClob) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public NClob getNClob(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public NClob getNClob(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public SQLXML getSQLXML(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public SQLXML getSQLXML(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void updateSQLXML(int columnIndex, SQLXML xmlObject) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateSQLXML(String columnLabel, SQLXML xmlObject) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public String getNString(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public String getNString(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Reader getNCharacterStream(int columnIndex) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Reader getNCharacterStream(String columnLabel) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void updateNCharacterStream(int columnIndex, Reader x, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNCharacterStream(String columnLabel, Reader reader, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateAsciiStream(int columnIndex, InputStream x, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBinaryStream(int columnIndex, InputStream x, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateCharacterStream(int columnIndex, Reader x, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateAsciiStream(String columnLabel, InputStream x, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBinaryStream(String columnLabel, InputStream x, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateCharacterStream(String columnLabel, Reader reader, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBlob(int columnIndex, InputStream inputStream, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBlob(String columnLabel, InputStream inputStream, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateClob(int columnIndex, Reader reader, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateClob(String columnLabel, Reader reader, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNClob(int columnIndex, Reader reader, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNClob(String columnLabel, Reader reader, long length) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNCharacterStream(int columnIndex, Reader x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNCharacterStream(String columnLabel, Reader reader) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateAsciiStream(int columnIndex, InputStream x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBinaryStream(int columnIndex, InputStream x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateCharacterStream(int columnIndex, Reader x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateAsciiStream(String columnLabel, InputStream x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBinaryStream(String columnLabel, InputStream x) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateCharacterStream(String columnLabel, Reader reader) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBlob(int columnIndex, InputStream inputStream) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateBlob(String columnLabel, InputStream inputStream) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateClob(int columnIndex, Reader reader) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateClob(String columnLabel, Reader reader) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNClob(int columnIndex, Reader reader) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void updateNClob(String columnLabel, Reader reader) throws SQLException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public <T> T getObject(int columnIndex, Class<T> type) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public <T> T getObject(String columnLabel, Class<T> type) throws SQLException {
		// TODO Auto-generated method stub
		return null;
	}
}