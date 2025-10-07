import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.net.http.*;
import java.net.URI;
import java.nio.file.Files;
import java.util.UUID;

public class RasTivClient extends JFrame {
    private final JTextField serverField = new JTextField("http://127.0.0.1:8000");
    private final JTextField fileField   = new JTextField();
    private File selected;

    public RasTivClient() {
        super("RAS/TIV Matrix Builder");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(600, 160);
        setLayout(new BorderLayout(6,6));

        JPanel top = new JPanel(new GridLayout(2,1,6,6));
        JPanel row1 = new JPanel(new BorderLayout(6,6));
        row1.add(new JLabel("Server: "), BorderLayout.WEST);
        row1.add(serverField, BorderLayout.CENTER);
        top.add(row1);

        JPanel row2 = new JPanel(new BorderLayout(6,6));
        row2.add(new JLabel("Excel file: "), BorderLayout.WEST);
        row2.add(fileField, BorderLayout.CENTER);
        JButton browse = new JButton("Browse");
        browse.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                selected = fc.getSelectedFile();
                fileField.setText(selected.getAbsolutePath());
            }
        });
        row2.add(browse, BorderLayout.EAST);
        top.add(row2);
        add(top, BorderLayout.NORTH);

        JPanel buttons = new JPanel();
        JButton rasBtn = new JButton("Build RAS");
        JButton tivBtn = new JButton("Build TIV");
        rasBtn.addActionListener(e -> submit("ras"));
        tivBtn.addActionListener(e -> submit("tiv"));
        buttons.add(rasBtn); buttons.add(tivBtn);
        add(buttons, BorderLayout.SOUTH);
    }

    private void submit(String mode) {
        if (selected == null) {
            JOptionPane.showMessageDialog(this, "Pick an .xlsx file first.");
            return;
        }
        try {
            String server = serverField.getText().trim();
            HttpClient client = HttpClient.newHttpClient();
            String boundary = "----" + UUID.randomUUID();
            // Build multipart body
            var baos = new ByteArrayOutputStream();
            try (var out = new DataOutputStream(baos)) {
                // file part
                out.writeBytes("--" + boundary + "\r\n");
                out.writeBytes("Content-Disposition: form-data; name=\"file\"; filename=\"" + selected.getName() + "\"\r\n");
                out.writeBytes("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n");
                out.write(Files.readAllBytes(selected.toPath()));
                out.writeBytes("\r\n");
                // end
                out.writeBytes("--" + boundary + "--\r\n");
            }
            HttpRequest req = HttpRequest.newBuilder()
                    .uri(URI.create(server + "/build?mode=" + mode))
                    .header("Content-Type", "multipart/form-data; boundary=" + boundary)
                    .POST(HttpRequest.BodyPublishers.ofByteArray(baos.toByteArray()))
                    .build();

            HttpResponse<byte[]> resp = client.send(req, HttpResponse.BodyHandlers.ofByteArray());
            if (resp.statusCode() != 200) {
                JOptionPane.showMessageDialog(this, "Server error: " + resp.statusCode() + "\n" + new String(resp.body()));
                return;
            }

            // Try to read filename from header, else synthesize
            String filename = "Output.xlsx";
            var disp = resp.headers().firstValue("Content-Disposition").orElse("");
            int idx = disp.indexOf("filename=");
            if (idx >= 0) filename = disp.substring(idx + 9).replace("\"", "").trim();

            File saveTo = new File(selected.getParentFile(), filename);
            Files.write(saveTo.toPath(), resp.body());
            JOptionPane.showMessageDialog(this, "Saved:\n" + saveTo.getAbsolutePath());
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(this, "Failed: " + ex.getMessage());
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new RasTivClient().setVisible(true));
    }
}
