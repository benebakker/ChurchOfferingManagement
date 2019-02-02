import java.util.Comparator;

public class CustomComparator implements Comparator<Donor> {
    @Override
    public int compare(Donor d1, Donor d2) {
    		System.out.println("in the custom comparator");
        return d1.getLastName().toUpperCase().compareTo(d2.getLastName().toUpperCase());
    }
}