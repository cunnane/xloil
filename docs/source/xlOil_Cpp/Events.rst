================
xlOil C++ Events
================

See :ref:`Events:Introduction`.

Examples
--------

.. highlight:: c++

::

    // When the returned shared_ptr is destroyed, the handler is unhooked.
    auto ptr = xloil::Event::CalcCancelled().bind(
          [this]() { this->cancel(); }));

    // The returned id can be used to unhook the handler
    static auto id = Event::AfterCalculate() += [logger]() { logger->flush(); };

