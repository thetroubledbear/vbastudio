// VbaStudio.Core/Dap/DapRequestReader.cs
using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;

namespace VbaStudio.Core.Dap;

/// <summary>
/// Single dedicated reader of the DAP input stream for the whole session. Every consumer (the
/// pre-launch request loop, OnProbe's paused read loop) must go through <see cref="Read"/> instead
/// of touching the stream directly - two independent readers racing on the same stream would tear
/// a Content-Length-framed message in half.
///
/// Reading always happens on this reader's own background thread. That is what makes a stop
/// request effective even while the session's main thread is blocked deep inside a synchronous
/// Application.Run COM call with nothing else pumping stdin: "disconnect"/"terminate" arriving
/// while <c>isPaused</c> reports false is handled immediately, inline, on this thread via
/// <c>handleWhileRunning</c> - instead of being queued for a consumer that will not call
/// <see cref="Read"/> again until that blocking call returns on its own.
/// </summary>
public sealed class DapRequestReader
{
    private readonly Stream _input;
    private readonly Func<bool> _isPaused;
    private readonly Action<DapRequest> _handleWhileRunning;
    private readonly BlockingCollection<DapRequest> _queue = new();
    private Thread? _thread;

    public DapRequestReader(Stream input, Func<bool> isPaused, Action<DapRequest> handleWhileRunning)
    {
        _input = input;
        _isPaused = isPaused;
        _handleWhileRunning = handleWhileRunning;
    }

    public void Start()
    {
        _thread = new Thread(ReadLoop) { IsBackground = true };
        _thread.Start();
    }

    /// <summary>
    /// Blocking dequeue. Returns null once the input stream has hit EOF/malformed input and every
    /// already-queued request has been consumed - mirrors DapProtocol.ReadRequest's own
    /// null-on-disconnect contract so existing callers need no further change in shape.
    /// </summary>
    public DapRequest? Read()
    {
        try
        {
            return _queue.Take();
        }
        catch (InvalidOperationException)
        {
            // BlockingCollection.Take() throws once CompleteAdding() has been called and the
            // queue is empty - this is the normal "stream closed, nothing left to read" path.
            return null;
        }
    }

    private void ReadLoop()
    {
        while (true)
        {
            var request = DapProtocol.ReadRequest(_input);
            if (request == null)
            {
                _queue.CompleteAdding();
                return;
            }

            var isStopCommand = request.Command == "disconnect" || request.Command == "terminate";
            if (isStopCommand && !_isPaused())
            {
                // Nobody is going to call Read() again soon - whichever consumer would is stuck
                // inside a blocking Application.Run call that this exact request is trying to
                // interrupt. Handle it here, now, on the reader thread, rather than queuing it to
                // sit unread until the run finishes on its own.
                _handleWhileRunning(request);
                continue;
            }

            _queue.Add(request);
        }
    }
}
