package org.tarasov;

import org.springframework.context.support.ClassPathXmlApplicationContext;

import java.io.IOException;
import java.util.Scanner;

public class ApplicationStarter {
    public static void main(String[] args) throws IOException {
        ClassPathXmlApplicationContext classPathXmlApplicationContext = new ClassPathXmlApplicationContext("ApplicationContext.xml");
        Parser parser = (Parser) classPathXmlApplicationContext.getBean("parser");
        String path;
        Scanner scanner = new Scanner(System.in);
        System.out.println("$/ : Введите путь с файлами для обработки: ");
        path = scanner.nextLine();
        parser.setFilesForProcessingPath(path);
        parser.process();
    }
}
