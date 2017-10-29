/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package satelite01;

import java.util.ArrayList;
import javax.swing.table.DefaultTableModel;

/**
 *
 * @author mzavaleta
 */
public class SateliteTableModel extends DefaultTableModel {

    public SateliteTableModel() {
        super();
    }
    private ArrayList< Class> tipos = new ArrayList<Class>();

    @Override
    public boolean isCellEditable(int row, int column) {
        return false;
    }

    public void addColumn(String nombre, int tipo) {
        addColumn(nombre);
    }

    @Override
    public Class getColumnClass(int columnIndex) {
        return tipos.get(columnIndex);
    }
    public void addColumnClass(Class tipo)
    {
        tipos.add(tipo);
    }
}
