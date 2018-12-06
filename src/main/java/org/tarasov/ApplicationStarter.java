package org.tarasov;

import org.springframework.context.support.ClassPathXmlApplicationContext;

import java.io.IOException;
import java.util.Scanner;

public class ApplicationStarter {
    public static void main(String[] args) throws IOException {
        String pathToSave = args[0];
        String patternsPath = args[1];
        System.out.println("Путь для сохранения обработанных файлов: " + pathToSave);
        System.out.println("Путь шаблонов: " + patternsPath);
        System.out.println();

        System.out.println();
        ClassPathXmlApplicationContext classPathXmlApplicationContext = new ClassPathXmlApplicationContext("ApplicationContext.xml");
        ColorResolver colorResolver = (ColorResolver) classPathXmlApplicationContext.getBean("colorResolver");
        colorResolver.setPatternsPath(patternsPath);
        Parser parser = (Parser) classPathXmlApplicationContext.getBean("parser");
        parser.setPathToSave(pathToSave);
        parser.setPatternsPath(colorResolver.getPatternsPath());

        colorResolver.init();

        String path;
        Scanner scanner = new Scanner(System.in);
        System.out.println("$/ : Введите путь с файлами для обработки: ");
        path = scanner.nextLine();
        parser.setFilesForProcessingPath(path);
        parser.process();
    }
}
