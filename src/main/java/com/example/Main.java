package com.example;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
// import io.github.bonigarcia.wdm.WebDriverManager;

public class Main {
    public static void main(String[] args) {
        // Setup ChromeDriver path (download from https://chromedriver.chromium.org/downloads)
        // System.setProperty("webdriver.chrome.driver", "path/to/chromedriver.exe");

        // Alternatively, use WebDriverManager if dependencies are resolved
        // WebDriverManager.chromedriver().setup();

        // For now, assuming chromedriver is in PATH or set
        // Create a new instance of ChromeDriver
        WebDriver driver = new ChromeDriver();

        // Navigate to Google
        driver.get("https://www.google.com");

        // Print the title of the page
        System.out.println("Page title is: " + driver.getTitle());

        // Close the browser
        driver.quit();
    }
}