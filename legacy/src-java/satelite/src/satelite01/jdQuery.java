/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/*
 * jdQuery.java
 *
 * Created on 02/08/2010, 12:41:29 PM
 */

package satelite01;

/**
 *
 * @author mzavaleta
 */
public class jdQuery extends javax.swing.JDialog {

    public void setQuery(String query)
    {
        jtaQuery.setText(query);
    }
    /** Creates new form jdQuery */
    public jdQuery(java.awt.Frame parent, boolean modal) {
        super(parent, modal);
        initComponents();
    }

    /** This method is called from within the constructor to
     * initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is
     * always regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        jtaQuery = new javax.swing.JTextArea();
        jToolBar1 = new javax.swing.JToolBar();
        jbOut = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setTitle("Ver consulta");

        jtaQuery.setColumns(20);
        jtaQuery.setEditable(false);
        jtaQuery.setFont(new java.awt.Font("Courier New", 0, 11)); // NOI18N
        jtaQuery.setRows(5);
        jScrollPane1.setViewportView(jtaQuery);

        jToolBar1.setFloatable(false);
        jToolBar1.setRollover(true);

        jbOut.setIcon(new javax.swing.ImageIcon(getClass().getResource("/satelite01/icons/door_out.png"))); // NOI18N
        jbOut.setToolTipText("Cerrar");
        jbOut.setFocusable(false);
        jbOut.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        jbOut.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        jbOut.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbOutActionPerformed(evt);
            }
        });
        jToolBar1.add(jbOut);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jToolBar1, javax.swing.GroupLayout.DEFAULT_SIZE, 536, Short.MAX_VALUE)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 536, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jToolBar1, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 296, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jbOutActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jbOutActionPerformed
        // TODO add your handling code here:
        this.dispose();
    }//GEN-LAST:event_jbOutActionPerformed

    /**
    * @param args the command line arguments
    */

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JToolBar jToolBar1;
    private javax.swing.JButton jbOut;
    private javax.swing.JTextArea jtaQuery;
    // End of variables declaration//GEN-END:variables

}
