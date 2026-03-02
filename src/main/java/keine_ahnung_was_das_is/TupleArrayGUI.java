package keine_ahnung_was_das_is;

import org.main.Tuple;

import javax.swing.*;
import java.awt.*;

public class TupleArrayGUI extends JFrame {
    private Tuple[][] tupleArray;

    public TupleArrayGUI(Tuple[][] tupleArray) {
        this.tupleArray = tupleArray;

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(600, 600);
        setLocationRelativeTo(null);

        CoordinateSystemPanel panel = new CoordinateSystemPanel();
        add(panel);

        setVisible(true);
    }

    class CoordinateSystemPanel extends JPanel {
        @Override
        protected void paintComponent(Graphics g) {
            super.paintComponent(g);

            int xOffset = getWidth() / 2;
            int yOffset = getHeight() / 2;

            // Draw X and Y axes
            g.drawLine(0, yOffset, getWidth(), yOffset);
            g.drawLine(xOffset, 0, xOffset, getHeight());

            // Draw points from the Tuple array
            for (Tuple[] row : tupleArray) {
                for (Tuple tuple : row) {
                    int x = xOffset + (int) tuple.getX();
                    int y = yOffset - (int) tuple.getY();

                    g.setColor(Color.RED);
                    g.fillOval(x - 5, y - 5, 10, 10);
                }
            }
        }
    }
}
