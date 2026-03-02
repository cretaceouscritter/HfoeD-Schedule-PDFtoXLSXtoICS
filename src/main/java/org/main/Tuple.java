package org.main;

public class Tuple {
    private final float x;
    private final float y;

    public Tuple(float x, float y) {
        this.x = x;
        this.y = y;
    }

    public float getX() {
        return x;
    }

    public float getY() {
        return y;
    }
    public String toString() {
    	return "(" + x + ", " + y + ")";
    }
}
