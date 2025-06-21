import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;
import java.util.concurrent.*;

class Worker {
    String name;
    List<Integer> tasks;
    int totalTime;
    int busyTime;
    int idleTime;

    Worker(String n, List<Integer> t) {
        name = n;
        tasks = t;
    }
}

public class TaskProcessor {
    static final int DAY_HOURS = 8;
    static List<Worker> workers = new ArrayList<>();

    public static void main(String[] args) throws Exception {
        loadData("workers.xlsx");
        ExecutorService exec = Executors.newFixedThreadPool(workers.size());
        
        for (Worker w : workers) {
            exec.execute(() -> {
                int dayLeft = DAY_HOURS;
                while (dayLeft > 0 && !w.tasks.isEmpty()) {
                    int task = w.tasks.get(0);
                    int spend = Math.min(task, dayLeft);
                    w.busyTime += spend;
                    dayLeft -= spend;
                    task -= spend;
                    if (task <= 0) w.tasks.remove(0);
                    else w.tasks.set(0, task);
                }
                w.idleTime = DAY_HOURS - w.busyTime;
                w.totalTime = DAY_HOURS;
            });
        }
        
        exec.shutdown();
        exec.awaitTermination(1, TimeUnit.HOURS);
        saveResults("results.xlsx");
        printStats();
    }

    static void loadData(String file) throws Exception {
        Workbook wb = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = wb.getSheetAt(0);
        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;
            String name = row.getCell(0).getStringCellValue();
            String[] tasks = row.getCell(1).getStringCellValue().split(",");
            List<Integer> taskList = new ArrayList<>();
            for (String t : tasks) taskList.add(Integer.parseInt(t.trim()));
            workers.add(new Worker(name, taskList));
        }
        wb.close();
    }

    static void saveResults(String file) throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Results");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Total");
        header.createCell(2).setCellValue("Busy");
        header.createCell(3).setCellValue("Idle");
        header.createCell(4).setCellValue("Tasks Left");
        
        int i = 1;
        for (Worker w : workers) {
            Row row = sheet.createRow(i++);
            row.createCell(0).setCellValue(w.name);
            row.createCell(1).setCellValue(w.totalTime);
            row.createCell(2).setCellValue(w.busyTime);
            row.createCell(3).setCellValue(w.idleTime);
            row.createCell(4).setCellValue(String.join(",", w.tasks.stream().map(Object::toString).toArray(String[]::new)));
        }
        
        FileOutputStream out = new FileOutputStream(file);
        wb.write(out);
        wb.close();
        out.close();
    }

    static void printStats() {
        System.out.println("=== STATS ===");
        workers.forEach(w -> {
            double eff = (double)w.busyTime / w.totalTime * 100;
            System.out.printf("%s: %dh work, %dh busy, %dh idle (%.1f%% eff)%n",
                w.name, w.totalTime, w.busyTime, w.idleTime, eff);
        });
    }
}