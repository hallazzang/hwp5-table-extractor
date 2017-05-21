import zlib

class ZlibDecompressStream(object):
    def __init__(self, stream, wbits=15, chunk_size=4096):
        self._stream = stream
        self._decompressor = zlib.decompressobj(wbits)
        self.chunk_size = chunk_size
        self.buffer = b''

    def read(self, size):
        while len(self.buffer) < size and not self._decompressor.eof:
            chunk = self._decompressor.unconsumed_tail
            if not chunk:
                chunk = self._stream.read(self.chunk_size)
                if not chunk:
                    break

            self.buffer += self._decompressor.decompress(chunk, self.chunk_size)

        result = self.buffer[:size]
        self.buffer = self.buffer[size:]

        return result