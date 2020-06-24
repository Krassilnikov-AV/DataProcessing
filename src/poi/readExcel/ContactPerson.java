
package poi.readExcel;

/*
 * Класс для хранения
 * перснональной информации контактного лица
 */

public class ContactPerson {
  
        private String name; //ФИО
        private String address; //Адрес
        private String phoneNumber; //Номер телефона

    public String getName() {
        return name;
    }

    public String getAddress() {
        return address;
    }

    public String getPhoneNumber() {
        return phoneNumber;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public void setPhoneNumber(String phoneNumber) {
        this.phoneNumber = phoneNumber;
    }
      

    
}
