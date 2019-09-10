package lv.accenture.start.it;

public class Hitman {

	private String firstname;
	private String lastname;
	private double price;

	public String getFirstname() {
		return firstname;
	}

	public Hitman(String firstname, String lastname, double price) {
		this.firstname = firstname;
		this.lastname = lastname;
		this.price = price;
	}

	public void setFirstname(String firstname) {
		this.firstname = firstname;
	}

	public String getLastname() {
		return lastname;
	}

	public void setLastname(String lastname) {
		this.lastname = lastname;
	}

	public double getPrice() {
		return price;
	}

	public void setPrice(double price) {
		this.price = price;
	}

}
