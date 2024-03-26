package org.example.model;

public class Employee {
  private int id;
  private String name;
  private String lastName;

  public Employee(int id, String name, String lastName) {
    this.id = id;
    this.name = name;
    this.lastName = lastName;
  }

  public int getId() {
    return id;
  }

  public String getName() {
    return name;
  }

  public String getLastName() {
    return lastName;
  }
}
