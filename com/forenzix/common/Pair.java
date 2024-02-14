package com.forenzix.common;

public class Pair<K, V> {

    public final K key;
    public final V value;

    public Pair(K key, V value) {
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

    public static <K, V> Pair<K, V> of(K key, V value) {
        return new Pair<>(key, value);
    }
}
