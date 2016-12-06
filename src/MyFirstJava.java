
public class MyFirstJava {

	public static void main(String[] args) {
		int a = 20;
		int b = 15;
		int c = calGCM(a, b);
		System.out.println(c);
	}

	private static int calGCM(int a, int b) {
		while (a % b != 0) {   
            int temp = a % b;   
            a = b;   
            b = temp;
        }   
        return b;
	}
}
