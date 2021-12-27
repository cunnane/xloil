
#include <pybind11/pybind11.h>
#include <list>
#include <mutex>
#include <thread>

namespace py = pybind11;

namespace xloil
{
  namespace Python
  {

    class GarbageCollector
    {
    public:
      static GarbageCollector& instance()
      {
        static GarbageCollector gc;
        return gc;
      }

      void add(PyObject* obj)
      {
        bool trashFull = false;
        {
          std::scoped_lock(_lock);
          _trash.push_back(obj);
          trashFull = _trash.size() == 10;
        }
        if (trashFull)
          _flag.notify_one();
      }
    private:
      std::list<PyObject*> _trash;
      std::mutex _lock;
      std::condition_variable _flag;
      std::thread _gcThread;
      std::atomic<bool> _running;

      GarbageCollector()
      {
        _running = true;
        _gcThread = std::thread([gc = this] 
        {
          std::unique_lock<std::mutex> lock(gc->_lock);
          do 
          {
            // Need to acquire lock before calling wait(lock)
            lock.lock();
            gc->_flag.wait(lock);
            if (gc->_trash.empty())
              continue;

            decltype(_trash) chute;
            std::swap(chute, gc->_trash);
            lock.unlock();
            {
              py::gil_scoped_acquire getGil;
              for (auto obj : chute)
                Py_XDECREF(obj);
            }
            chute.clear();
          } while (gc->_running);
        });
      }
      ~GarbageCollector()
      {
        _running = false;
        _flag.notify_one();
        _gcThread.join();
      }
    };

    class objectGC : py::handle
    {
    public:
      //objectGC() = default;
      objectGC(py::handle h, bool is_borrowed) 
        : py::handle(h) 
      { 
        if (is_borrowed) inc_ref(); 
      }
      /// Copy constructor; always increases the reference count
      objectGC(const objectGC& o) : py::handle(o) { inc_ref(); }
      /// Move constructor; steals the object from ``other`` and preserves its reference count
      objectGC(objectGC&& other) noexcept { m_ptr = other.m_ptr; other.m_ptr = nullptr; }
      ~objectGC()
      {
        GarbageCollector::instance().add(m_ptr);
      }
    };
  }
}