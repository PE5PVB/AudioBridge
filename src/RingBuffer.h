#pragma once

#include <atomic>
#include <cstring>
#include <vector>
#include <algorithm>

// Single-Producer Single-Consumer lock-free ring buffer.
// Stores raw audio bytes. Thread-safe without mutexes.
class RingBuffer {
public:
    explicit RingBuffer(size_t capacityBytes)
        : m_buffer(capacityBytes > 0 ? capacityBytes : 1)
        , m_capacity(capacityBytes > 0 ? capacityBytes : 1)
        , m_head(0)
        , m_tail(0)
    {}

    void reset() {
        m_head.store(0, std::memory_order_relaxed);
        m_tail.store(0, std::memory_order_relaxed);
    }

    size_t capacity() const { return m_capacity; }

    size_t availableToRead() const {
        size_t h = m_head.load(std::memory_order_acquire);
        size_t t = m_tail.load(std::memory_order_relaxed);
        return (h >= t) ? (h - t) : (m_capacity - t + h);
    }

    size_t availableToWrite() const {
        return m_capacity - 1 - availableToRead();
    }

    // Producer: write data into the ring buffer.
    // Returns number of bytes actually written.
    size_t write(const void* data, size_t bytes) {
        size_t avail = availableToWrite();
        size_t toWrite = (std::min)(bytes, avail);
        if (toWrite == 0) return 0;

        size_t h = m_head.load(std::memory_order_relaxed);
        const uint8_t* src = static_cast<const uint8_t*>(data);

        size_t firstPart = (std::min)(toWrite, m_capacity - h);
        std::memcpy(m_buffer.data() + h, src, firstPart);
        if (toWrite > firstPart) {
            std::memcpy(m_buffer.data(), src + firstPart, toWrite - firstPart);
        }

        size_t newHead = (h + toWrite) % m_capacity;
        m_head.store(newHead, std::memory_order_release);
        return toWrite;
    }

    // Consumer: read data from the ring buffer.
    // Returns number of bytes actually read.
    size_t read(void* dest, size_t bytes) {
        size_t avail = availableToRead();
        size_t toRead = (std::min)(bytes, avail);
        if (toRead == 0) return 0;

        size_t t = m_tail.load(std::memory_order_relaxed);
        uint8_t* dst = static_cast<uint8_t*>(dest);

        size_t firstPart = (std::min)(toRead, m_capacity - t);
        std::memcpy(dst, m_buffer.data() + t, firstPart);
        if (toRead > firstPart) {
            std::memcpy(dst + firstPart, m_buffer.data(), toRead - firstPart);
        }

        size_t newTail = (t + toRead) % m_capacity;
        m_tail.store(newTail, std::memory_order_release);
        return toRead;
    }

private:
    std::vector<uint8_t> m_buffer;
    size_t               m_capacity;
    // Separate cache lines to avoid false sharing
    alignas(64) std::atomic<size_t> m_head;
    alignas(64) std::atomic<size_t> m_tail;
};
