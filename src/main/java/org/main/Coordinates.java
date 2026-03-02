package org.main;

public class Coordinates {
    final float X;
    final float Y;

    public Coordinates(float x, float y) {
        this.X = x;
        this.Y = y;
    }

    public String toString() {
    	return "(" + X + ", " + Y + ")";
    }
}
