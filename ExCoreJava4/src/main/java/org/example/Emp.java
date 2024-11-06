package org.example;

public class Emp {
    private static int index = 0;
    private int EMP_INDEX;
    private String ID;
    private String NAME;

    public Emp(){
        EMP_INDEX = index;
        index += 1;
    }

    public Emp(String id, String name){
        EMP_INDEX = index;
        ID = id;
        NAME = name;
        index += 1;
    }

    public Emp(Emp emp) {
        EMP_INDEX = index;
        ID = emp.getID();
        NAME = emp.getNAME();
        index += 1;
    }

    public void setID(String id){
        this.ID = id;
    }

    public void setNAME(String name){
        this.NAME = name;
    }

    public int getEMP_INDEX() {
        return EMP_INDEX;
    }

    public String getID(){
        return this.ID;
    }

    public String getNAME() {
        return NAME;
    }

}
