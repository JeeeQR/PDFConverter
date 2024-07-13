package org.pdf;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.dnd.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class OfficeToPDFConverterGUI extends JFrame {
    private DefaultListModel<File> listModel;
    private JList<File> fileList;
    private JTextField outputDirectoryField;
    private JProgressBar progressBar;
    private JTextArea statusArea;
    private File lastDirectory;

    public OfficeToPDFConverterGUI() {
        setTitle("PDF 转换器");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 800);
        setLocationRelativeTo(null);

        if (UIManager.getLookAndFeel().isSupportedLookAndFeel()) {
            final String platform = UIManager.getSystemLookAndFeelClassName();
            if (!UIManager.getLookAndFeel().getName().equals(platform)) {
                try {
                    UIManager.setLookAndFeel(platform);
                } catch (Exception exception) {
                    exception.printStackTrace();
                }
            }
        }

        JPanel contentPane = new JPanel(new BorderLayout());
        contentPane.setBorder(new EmptyBorder(10, 10, 10, 10));
        setContentPane(contentPane);

        listModel = new DefaultListModel<>();
        fileList = new JList<>(listModel);
        fileList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        JScrollPane fileListScrollPane = new JScrollPane(fileList);
        fileListScrollPane.setBorder(BorderFactory.createTitledBorder("要转化的文件"));
        contentPane.add(fileListScrollPane, BorderLayout.CENTER);

        JPanel outputPanel = new JPanel(new BorderLayout());
        outputPanel.setBorder(BorderFactory.createTitledBorder("PDF 输出目录"));
        outputDirectoryField = new JTextField();
        JButton browseButton = new JButton("选择输出目录");
        browseButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                if (lastDirectory != null) {
                    fileChooser.setCurrentDirectory(lastDirectory);
                }
                int option = fileChooser.showOpenDialog(OfficeToPDFConverterGUI.this);
                if (option == JFileChooser.APPROVE_OPTION) {
                    lastDirectory = fileChooser.getSelectedFile();
                    outputDirectoryField.setText(lastDirectory.getAbsolutePath());
                }
            }
        });
        outputPanel.add(outputDirectoryField, BorderLayout.CENTER);
        outputPanel.add(browseButton, BorderLayout.EAST);
        contentPane.add(outputPanel, BorderLayout.NORTH);

        JPanel statusPanel = new JPanel(new BorderLayout());
        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        statusPanel.add(progressBar, BorderLayout.NORTH);
        statusArea = new JTextArea(7, 20);
        statusArea.setFont(statusArea.getFont().deriveFont(Font.PLAIN, 16));
        statusArea.setEditable(false);
        JScrollPane statusScrollPane = new JScrollPane(statusArea);
        statusScrollPane.setBorder(BorderFactory.createTitledBorder("状态"));
        statusPanel.add(statusScrollPane, BorderLayout.CENTER);
        contentPane.add(statusPanel, BorderLayout.SOUTH);

        JPanel buttonPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        JButton addButton = new JButton("添加文件");
        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setMultiSelectionEnabled(true);
                if (lastDirectory != null) {
                    fileChooser.setCurrentDirectory(lastDirectory);
                }
                int option = fileChooser.showOpenDialog(OfficeToPDFConverterGUI.this);
                if (option == JFileChooser.APPROVE_OPTION) {
                    File[] selectedFiles = fileChooser.getSelectedFiles();
                    for (File file : selectedFiles) {
                        listModel.addElement(file);
                    }
                    lastDirectory = fileChooser.getCurrentDirectory();
                }
            }
        });
        gbc.gridx = 0;
        gbc.gridy = 0;
        buttonPanel.add(addButton, gbc);

        JButton convertButton = new JButton("转换成 PDF");
        convertButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                new Thread(() -> convertFilesToPDF()).start();
            }
        });
        gbc.gridx = 1;
        gbc.gridy = 0;
        buttonPanel.add(convertButton, gbc);

        JButton clearButton = new JButton("清空文件列表");
        clearButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                listModel.clear();
                statusArea.setText("");
            }
        });
        gbc.gridx = 0;
        gbc.gridy = 1;
        buttonPanel.add(clearButton, gbc);

        JButton openButton = new JButton("打开输出目录");
        openButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    Desktop.getDesktop().open(new File(outputDirectoryField.getText()));
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });
        gbc.gridx = 1;
        gbc.gridy = 1;
        buttonPanel.add(openButton, gbc);

        contentPane.add(buttonPanel, BorderLayout.EAST);

        new FileDropTarget(fileList);

        setVisible(true);
    }

    private void convertFilesToPDF() {
        List<File> files = new ArrayList<>();
        for (int i = 0; i < listModel.size(); i++) {
            files.add(listModel.getElementAt(i));
        }
        String outputDir = outputDirectoryField.getText();
        if (outputDir.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请选择一个具体的输出目录！", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        if (files.size() == 1) {
            progressBar.setIndeterminate(true); // 设置为不确定模式
        } else {
            progressBar.setIndeterminate(false);
            progressBar.setMaximum(files.size());
            progressBar.setValue(0);
        }
        statusArea.setText("");

        for (int i = 0; i < files.size(); i++) {
            File file = files.get(i);
            statusArea.append("正在转换 " + file.getName() + "...\n");
            if (files.size() > 1) {
                progressBar.setValue(i + 1);
            }

            try {
                convertOfficeToPDF(file, outputDir);
                statusArea.append("转换 " + file.getName() + " 成功.\n");
            } catch (Exception e) {
                statusArea.append("转换失败 " + file.getName() + ": " + e.getMessage() + "\n");
            }
        }

        if (files.size() == 1) {
            progressBar.setIndeterminate(false); // 恢复正常模式
        }

        statusArea.append("转换结束.\n");
    }

    private void convertOfficeToPDF(File file, String outputDir) throws IOException {

        String fileName = file.getName();
        String fileExtension = fileName.substring(fileName.lastIndexOf('.') + 1);

        String outputFilePath = outputDir + File.separator + fileName.substring(0, fileName.lastIndexOf('.')) + ".pdf";

        boolean conversionResult = false;

        switch (fileExtension.toLowerCase()) {
            case "doc":
            case "docx":
                conversionResult = APIS.doc2pdf(file.getAbsolutePath(), outputFilePath);
                break;
            case "xls":
            case "xlsx":
                conversionResult = APIS.xlsx2Pdf(file.getAbsolutePath(), outputFilePath);
                break;
            case "ppt":
            case "pptx":
                conversionResult = APIS.ppt2Pdf(file.getAbsolutePath(), outputFilePath);
                break;
            case "txt":
                conversionResult = APIS.txt2pdf(file.getAbsolutePath(), outputFilePath);
                break;
            default:
                throw new IOException("Unsupported file type: " + fileExtension);
        }

        if (!conversionResult) {
            throw new IOException("Failed to convert " + fileName + " to PDF.");
        }
    }

    private class FileDropTarget extends DropTarget {
        public FileDropTarget(JList<File> fileList) {
            new DropTarget(fileList, DnDConstants.ACTION_COPY, new DropTargetAdapter() {
                @Override
                public void drop(DropTargetDropEvent dtde) {
                    try {
                        dtde.acceptDrop(DnDConstants.ACTION_COPY);
                        Transferable transferable = dtde.getTransferable();
                        List<File> droppedFiles = (List<File>) transferable.getTransferData(DataFlavor.javaFileListFlavor);
                        for (File file : droppedFiles) {
                            listModel.addElement(file);
                        }
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }
            }, true);
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new OfficeToPDFConverterGUI());
    }
}
