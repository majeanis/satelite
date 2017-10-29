/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package model;

import java.awt.Component;
import javax.swing.ImageIcon;
import javax.swing.JTree;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeCellRenderer;
import satelite01.MdiMain;

/**
 *
 * @author mzavaleta
 */
public class JTreeIconRender extends DefaultTreeCellRenderer {

    ImageIcon mainGrpIcon;
    ImageIcon grupoIcon;
    ImageIcon queryIcon;
    ImageIcon MBdIcon;
    ImageIcon MQryIcon;

    public JTreeIconRender() {
        mainGrpIcon= new ImageIcon(JTreeIconRender.class.getResource("/satelite01/icons/folder_table.png"));
        MQryIcon = new ImageIcon(JTreeIconRender.class.getResource("/satelite01/icons/report_link.png"));
        MBdIcon = new ImageIcon(JTreeIconRender.class.getResource("/satelite01/icons/database_link.png"));
        grupoIcon = new ImageIcon(JTreeIconRender.class.getResource("/satelite01/icons/group.png"));
        queryIcon = new ImageIcon(JTreeIconRender.class.getResource("/satelite01/icons/report_go.png"));
        //group.png
        //report_go.png
    }

    @Override
    public Component getTreeCellRendererComponent(JTree tree,
            Object value, boolean sel, boolean expanded, boolean leaf,
            int row, boolean hasFocus) {

        super.getTreeCellRendererComponent(tree, value, sel,
                expanded, leaf, row, hasFocus);

        Object nodeObj = ((DefaultMutableTreeNode) value).getUserObject();
        if (nodeObj.getClass().getName().endsWith("String")) {
            if (((String) nodeObj).equals(MdiMain.cTEXT_MTTO_CONSULTAS)) {
                setIcon(MQryIcon);
            } else if (((String) nodeObj).equals(MdiMain.cTEXT_MTTO_BDATOS)) {
                setIcon(MBdIcon);
            }else if (((String) nodeObj).equals(MdiMain.cTEXT_CONSULTAS)) {
                setIcon(mainGrpIcon);
            }
        }else if(nodeObj.getClass().getName().endsWith("BE_Grupo")){
            setIcon(grupoIcon);
        }else if(nodeObj.getClass().getName().endsWith("BE_Consulta")){
            setIcon(queryIcon);
        }
        return this;
    }
}
