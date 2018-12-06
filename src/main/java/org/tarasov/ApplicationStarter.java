package org.tarasov;

import org.springframework.context.support.ClassPathXmlApplicationContext;

import java.io.IOException;
import java.util.Scanner;

public class ApplicationStarter {
    public static void main(String[] args) throws IOException {
        boolean done = false;
        while (!done) {
            Scanner scanner = new Scanner(System.in);
            String pathToSave = args[0];
            String patternsPath = args[1];
            System.out.println("Путь для сохранения обработанных файлов: " + pathToSave);
            System.out.println("Путь шаблонов: " + patternsPath);
            System.out.println();
            String pathFrom = args[2];
            String answer;
            System.out.println("По умолчанию?");
            answer = scanner.nextLine();
            System.out.println();

            if (answer.equals("y")) {
                pathFrom = args[2];
            } else if (answer.equals("n")) {
                System.out.println("$/ : Введите путь с файлами для обработки: ");
                pathFrom = scanner.nextLine();
            } else {
                System.out.println("Некорректный вариант ответа!");
                System.out.println();
                continue;
            }

            ClassPathXmlApplicationContext classPathXmlApplicationContext = new ClassPathXmlApplicationContext("ApplicationContext.xml");
            ColorResolver colorResolver = (ColorResolver) classPathXmlApplicationContext.getBean("colorResolver");
            colorResolver.setPatternsPath(patternsPath);
            Parser parser = (Parser) classPathXmlApplicationContext.getBean("parser");
            parser.setPathToSave(pathToSave);
            parser.setPatternsPath(colorResolver.getPatternsPath());

            colorResolver.init();
            parser.setFilesForProcessingPath(pathFrom);
            parser.process();
            done = true;
        }
    }
}
