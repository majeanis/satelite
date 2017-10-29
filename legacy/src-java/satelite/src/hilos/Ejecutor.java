/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package hilos;


import BE.LBE_Consulta;
import BU.Ex.LBU_Exception;
import BU.LBU_Consulta;
import java.awt.Component;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Vector;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableColumnModel;
import logger.Log;
import satelite01.Constantes;
import satelite01.SateliteTableModel;
import satelite01.jifConsulta;

/**
 *
 * @author mzavaleta
 */
public class Ejecutor implements Runnable {

    private LBE_Consulta consulta;
    private String query;
    private jifConsulta form;

    public Ejecutor(int id, String query, jifConsulta form) {
        consulta = new LBE_Consulta();
        this.query = query;
        consulta.setId(id);
        this.form = form;
    }

    public void run() {
        LBU_Consulta bu_Consulta = null;
        this.form.setInicio(new Date());
        long contador = 0;
        try {
            this.form.jProgressBar1.setIndeterminate(true);
            this.form.jProgressBar1.setString("Ejecutando consulta..");
            bu_Consulta = new LBU_Consulta(Constantes.UsuarioLogin.getBDatos());
            form.marcaInicio("");
            consulta = bu_Consulta.get(consulta.getId());
            Log.debug("Id de consulta " + consulta.getId());
            ResultSet rs = bu_Consulta.ejecutar(consulta, query, Constantes.UsuarioLogin.getUsrLogin(), false);
            //rs.setFetchSize(5);
            Log.debug("FetchSize: " + rs.getFetchSize());
            int id = consulta.getId();

            String arrWidthCols = Constantes.propiedades.getProperty("consulta." + String.valueOf(id) + ".cols", "75");
            ArrayList<Integer> tiposCol = new ArrayList<Integer>();
            //Se agrega las columnas
            SateliteTableModel model = (SateliteTableModel) form.jTable1.getModel();
            model.getDataVector().removeAllElements();
            model.setColumnCount(0);
            int numCols = rs.getMetaData().getColumnCount();
            //Class[] types = new Class[]{numCols };
            for (int ix = 1; ix <= numCols; ix++) {
                int columnType = rs.getMetaData().getColumnType(ix);
                String columnName = rs.getMetaData().getColumnName(ix);

                tiposCol.add(columnType);
                if (columnType == Types.DOUBLE || columnType == Types.NUMERIC
                        || columnType == Types.DECIMAL) {
                    model.addColumnClass(java.lang.Double.class);
                } else if (columnType == Types.INTEGER || columnType == Types.BIGINT) {
                    model.addColumnClass(java.lang.Integer.class);
                } else if (columnType == Types.DATE) {
                    model.addColumnClass(java.util.Date.class);
                } else {
                    model.addColumnClass(java.lang.String.class);
                }
                model.addColumn(columnName);
            }

            while (rs.next()) {
                contador++;
                if (contador % 1000 == 0) {
                    this.form.jlMensaje.setText("Van " + Long.toString(contador) + " registros");
                }

                Vector row = new Vector();
                for (int ix = 1; ix <= numCols; ix++) {
                    int columnType = tiposCol.get(ix - 1);

                    switch (columnType) {
                        case Types.DOUBLE:
                            row.add(rs.getDouble(ix));
                            break;
                        case Types.NUMERIC:
                            row.add(rs.getDouble(ix));
                            break;
                        default:
                            row.add(rs.getObject(ix));
                            break;
                    }
                }


                model.addRow(row);
            }
            String[] widths = arrWidthCols.split(",");
            DefaultTableColumnModel modelColumn = (DefaultTableColumnModel) form.jTable1.getColumnModel();
            for (int ix = 0; ix < modelColumn.getColumnCount(); ix++) {
                if (widths.length > ix) {
                    int width = Integer.parseInt(widths[ix]);
                    modelColumn.getColumn(ix).setPreferredWidth(width);
                }
            }
            this.form.setFinal(new Date());
            this.form.jProgressBar1.setString("100%");
            form.marcaFin("Consulta generada en " + String.valueOf(this.form.deltaTiempo()) + " segs, "
                    + String.valueOf(contador) + " regs. generados");
        } catch (LBU_Exception ex) {
            this.form.jProgressBar1.setString("Error");
            form.marcaError("Error en consulta: " + ex.getMessage());
        } catch (SQLException ex) {
            this.form.jProgressBar1.setString("Error");
            form.marcaError("Error en consulta: " + ex.getMessage());
        } catch (NullPointerException ex) {
            this.form.jProgressBar1.setString("Error");
            Log.error("Error en consulta", ex);
            form.marcaError("Error en consulta: Favor reintentar");
        } catch (ArrayIndexOutOfBoundsException ex) {
            this.form.jProgressBar1.setString("Error");
            Log.error("Error en consulta", ex);
            form.marcaError("Error en consulta: Favor reintentar");
        } catch (Exception ex) {
            this.form.jProgressBar1.setString("Error");
            Log.error("Error en consulta, id:" + consulta.getId() , ex);
            form.marcaError("Error en consulta: " + ex.getMessage());
        } catch (OutOfMemoryError err) {
            this.form.jProgressBar1.setString("Error");
            //Log.;
            Log.error("Error en consulta ( Registro numero " + contador + "), consulta id :" + consulta.getId(), err);
            form.marcaError("Error en consulta, por falta de memoria : " + err.getMessage());
        } finally {
            this.form.jProgressBar1.setIndeterminate(false);
            this.form.jProgressBar1.setMaximum(1);
            this.form.jProgressBar1.setValue(1);
        }
    }

    static class DecimalFormatRenderer extends DefaultTableCellRenderer {

        private static final DecimalFormat formatter = new DecimalFormat("#.00");

        @Override
        public Component getTableCellRendererComponent(
                JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
            value = formatter.format(value);
            return super.getTableCellRendererComponent(
                    table, value, isSelected, hasFocus, row, column);
        }
    }
}
