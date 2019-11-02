package com.iropensource.model;

public final class User {
    private String firstName;
    private String lastName;
    private String username;
    private String emailAddress;

    public User(String firstName, String lastName, String username, String emailAddress) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.username = username;
        this.emailAddress = emailAddress;
    }

    public String getFirstName() {
        return firstName;
    }

    public String getLastName() {
        return lastName;
    }

    public String getUsername() {
        return username;
    }

    public String getEmailAddress() {
        return emailAddress;
    }

    @Override
    public String toString() {
        return String.format("First Name: %s | Last Name: %s | Username: %s | Email Address: %s", firstName, lastName, username, emailAddress);
    }
}