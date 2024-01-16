package com.forenzix.common;

public class Slot<V> {
    
    public final String key;
    private V value;

    public Slot(String key, V value) {
        if (key == null) {
            throw new IllegalArgumentException("Key cannot be null.");
        }

        this.key = key;
        this.value = value;
    }

    @Override
    public String toString() {
        return String.format("{ \"%s\" : \"%s\" }", key, value);
    }

    public static <V> Slot<V> of(String key, V value) {
        return new Slot<>(key, value);
    }

    public V value() {
        return value;
    }

    public Slot<V> value(V newValue) {
        this.value = newValue;
        return this;
    }
}
