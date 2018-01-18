class BitDemo {
    public static void main(String[] args) {
        int bitmask = 0x0F00;
        int val = 0x22FF;
        // prints "2"
        System.out.println(val);
        System.out.println("Bit And: " + (val & bitmask));
        System.out.println("Bit And: " + (val ^ bitmask));
    }
}