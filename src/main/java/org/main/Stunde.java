package org.main;

public class Stunde {
	private boolean belegt = false;
    private String inhalt;
    private String raum;
    private String lehrkraft;
    public String getInhalt() {
        return inhalt;
    }

    public void setInhalt(String inhalt) {
        this.inhalt = inhalt;
    }

    public String getRaum() {
        return raum;
    }

    public void setRaum(String raum) {
        this.raum = raum;
    }

    public String getLehrkraft() {
        return lehrkraft;
    }

    public void setLehrkraft(String lehrkraft) {
        this.lehrkraft = lehrkraft;
    }
    public void setBelegt() {
        this.belegt = true;
    }
    public boolean istBelegt() {
    	return belegt;
    }
   
    @Override
    public String toString() {
    	if(!belegt)return "leere Stunde";
        return "Inhalt: " + inhalt + ", Lehrkraft: " + lehrkraft + ", Raum: " + raum;
    }
}