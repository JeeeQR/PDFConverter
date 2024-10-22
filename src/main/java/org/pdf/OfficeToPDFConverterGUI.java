package org.pdf;

import org.pdf.APIS;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.dnd.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;
import java.util.List;
import java.util.concurrent.*;

public class OfficeToPDFConverterGUI extends JFrame {

    private final DefaultListModel<File> listModel = new DefaultListModel<>();
    private final JTextField outputDirectoryField;
    private final JProgressBar progressBar;
    private final JTextArea statusArea;
    private File lastDirectory;

    private void loadSettings() {
        File settingsDirectory = new File(System.getProperty("user.home"), ".pdfconverter");
        File settingsFile = new File(settingsDirectory, "settings.properties");
        if (!settingsFile.exists()) {
            createDefaultSettings(settingsFile);
        } else {
            Properties properties = new Properties();
            try (FileInputStream fis = new FileInputStream(settingsFile)) {
                properties.load(fis);
                String outputPath = properties.getProperty("outputPath");
                if (!outputPath.isEmpty()) {
                    outputDirectoryField.setText(outputPath);
                }
                String core = properties.getProperty("core");
                if (core == null || core.isEmpty()) {
                    properties.setProperty("core", "1");
                    try (FileOutputStream fos = new FileOutputStream(settingsFile)) {
                        properties.store(fos, "Updated settings with core property");
                    } catch (IOException e) {
                        System.err.println("Error updating settings file: " + e.getMessage());
                    }
                }
            } catch (IOException e) {
                System.err.println("Error loading settings from file: " + e.getMessage());
            }
        }
    }

    private void createDefaultSettings(File settingsFile) {
        File settingsDirectory = new File(System.getProperty("user.home"), ".pdfconverter"); // 应用数据目录为用户主目录下的 .pdfconverter 文件夹
        settingsDirectory.mkdirs(); // 如果目录不存在，则创建它
        if (!settingsFile.exists()) { // 只有当文件不存在时才创建
            Properties properties = new Properties();
            properties.setProperty("outputPath", System.getProperty("user.home") + "/Documents");
            properties.setProperty("core", "1"); // 设置转换用核心数
            try {
                properties.store(new FileOutputStream(settingsFile), "Default application settings");
            } catch (IOException e) {
                System.out.println("Error creating default settings file: " + e.getMessage());
            }
        }
    }

    private void saveSettings() {
        File settingsDirectory = new File(System.getProperty("user.home"), ".pdfconverter"); // 应用数据目录为用户主目录下的 .pdfconverter 文件夹
        File settingsFile = new File(settingsDirectory, "settings.properties"); // 设置文件名
        Properties properties = new Properties();
        properties.setProperty("outputPath", outputDirectoryField.getText());
        properties.setProperty("core", "1"); // 设置转换用核心数
        try (FileOutputStream fos = new FileOutputStream(settingsFile)) {
            properties.store(fos, "Application settings");
        } catch (IOException e) {
            System.out.println("Error saving settings to file: " + e.getMessage());
        }
    }

    private void convertFilesToPDF() {
        ArrayList<File> files = new ArrayList<>(listModel.size()); // 创建一个大小与 listModel 相同的新 ArrayList
        for (int i = 0; i < listModel.getSize(); i++) {
            files.add((File) listModel.getElementAt(i)); // 将 listModel 中的每个元素添加到新 ArrayList 中
        }
        String outputDir = outputDirectoryField.getText();
        if (outputDir.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请输入一个具体的输出目录！", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }
        // 获取最大线程数
        int maxThreads = APIS.getCoreValue();
        ExecutorService executor = Executors.newFixedThreadPool(maxThreads);
        progressBar.setIndeterminate(false);
        progressBar.setMaximum(files.size());
        progressBar.setValue(0);
        statusArea.setText("");
        for (int i = 0; i < files.size(); i++) {
            final File file = files.get(i);
            int finalI = i;
            Runnable task = new Runnable() {
                @Override
                public void run() {
                    statusArea.append("正在转换 " + file.getName() + "...\n");
                    try {
                        convertOfficeToPDF(file, outputDir);
                        SwingUtilities.invokeLater(new Runnable() {
                            @Override
                            public void run() {
                                statusArea.append("转换 " + file.getName() + " 成功。\n");
                            }
                        });
                    } catch (Exception e) {
                        SwingUtilities.invokeLater(new Runnable() {
                            @Override
                            public void run() {
                                statusArea.append("转换失败 " + file.getName() + ": " + e.getMessage() + "\n");
                            }
                        });
                    } finally {
                        SwingUtilities.invokeLater(new Runnable() {
                            @Override
                            public void run() {
                                progressBar.setValue(finalI + 1);
                            }
                        });
                    }
                }
            };
            executor.execute(task);
        }
        executor.shutdown();
        try {
            executor.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    private void convertOfficeToPDF(File file, String outputDir) throws Exception {
        String outputFilePath = outputDir + "/" + file.getName().replaceFirst("[.][^.]+$", "") + ".pdf";
        String fileName = file.getName();
        String fileExtension = fileName.substring(fileName.lastIndexOf('.') + 1);

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
                throw new IllegalArgumentException("不支持的文件格式：" + fileExtension);
        }
        if (!conversionResult) {
            throw new IOException("Failed to convert " + fileName + " to PDF.");
        }
    }

    private class FileDropTarget extends DropTarget {
        private final OfficeToPDFConverterGUI parent;

        public FileDropTarget(JList<File> fileList, OfficeToPDFConverterGUI parent) {
            super(fileList, DnDConstants.ACTION_COPY, new DropTargetAdapter() {
                @Override
                public void drop(DropTargetDropEvent dtde) {
                    try {
                        dtde.acceptDrop(DnDConstants.ACTION_COPY);
                        Transferable transferable = dtde.getTransferable();
                        if (transferable.isDataFlavorSupported(DataFlavor.javaFileListFlavor)) {
                            List<?> droppedFiles = (List<?>) transferable.getTransferData(DataFlavor.javaFileListFlavor);
                            for (Object fileObj : droppedFiles) {
                                if (fileObj instanceof File) {
                                    File file = (File) fileObj;
                                    SwingUtilities.invokeLater(() -> parent.listModel.addElement(file));
                                }
                            }
                        }
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }
            }, true);

            this.parent = parent; // 设置 parent 变量的值
        }
    }

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

        // 创建一个JPanel，使用边框布局
        JPanel contentPane = new JPanel(new BorderLayout());
        contentPane.setBorder(new EmptyBorder(10, 10, 10, 10));
        setContentPane(contentPane);

        JList<File> fileList = new JList<>(listModel); // 传递 listModel 而不是它的 toArray 结果
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
                    saveSettings();
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
                        listModel.addElement(file); // 使用 addElement 方法逐个添加文件
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
                SwingUtilities.invokeLater(new Runnable() { // 确保在 EDT 中更新模型
                    @Override
                    public void run() {
                        listModel.clear();
                        statusArea.setText("");
                    }
                });
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

        new FileDropTarget(fileList, this);
        loadSettings();
        setVisible(true);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(OfficeToPDFConverterGUI::new);
    }
}