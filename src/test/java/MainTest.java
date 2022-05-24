import org.openqa.selenium.chrome.ChromeDriver;

public class MainTest {
    public static void test(){
        //
        ChromeDriver dr=new ChromeDriver();
        dr.get("https://www.baidu.com");
    }

    public static void main(String[] args) {
        test();
    }


}