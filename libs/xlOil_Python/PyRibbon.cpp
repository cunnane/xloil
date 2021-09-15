#include "Main.h"
#include "BasicTypes.h"
#include "PyCore.h"
#include "PyEvents.h"
#include "PyImage.h"
#include <xloil/Ribbon.h>
#include <xloil/RtdServer.h>
#include <pybind11/pybind11.h>
#include <filesystem>
namespace py = pybind11;
using std::shared_ptr;
using std::wstring;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      auto setRibbon(IComAddin* addin, const wchar_t* xml, py::object mapper)
      {
        if (PyDict_Check(mapper.ptr()))
          mapper = mapper.attr("__getitem__");

        auto cmapper = [mapper](const wchar_t* name) // PyObjectHolder
        {
          try
          {
            py::gil_scoped_acquire getGil;
            auto callback = mapper(name);
            return [callback](
              const RibbonControl& ctrl, VARIANT* vRet, int nArgs, VARIANT** vArgs)
            {
              try
              {
                // Converting via an ExcelObj is not optimal, but avoids building
                // a VARIANT converter. Since very few VARIANT types can appear in
                // callbacks this might be a feasible approach when IPictureDisp
                // support is introduced.
                py::gil_scoped_acquire getGil;
                py::tuple args(nArgs);
                for (auto i = 0; i < nArgs; ++i)
                  args[i] = PyFromAny()(variantToExcelObj(*vArgs[i]));
                auto pyRet = callback(ctrl, *args);
                if (vRet && !pyRet.is_none())
                {
                  auto picture = pictureFromPilImage(pyRet);
                  if (picture)
                  {
                    VariantInit(vRet);
                    vRet->pdispVal = (IDispatch*)picture;
                    vRet->vt = VT_DISPATCH;
                  }
                  else
                    excelObjToVariant(vRet, FromPyObj<false>()(pyRet.ptr()));
                }
              }
              catch (const py::error_already_set& e)
              {
                Event_PyUserException().fire(e.type(), e.value(), e.trace());
                throw;
              }
            };
          }
          catch (const py::error_already_set& e)
          {
            Event_PyUserException().fire(e.type(), e.value(), e.trace());
            throw;
          }
        };
        addin->setRibbon(xml, cmapper);
      }
      auto createRibbon(const wchar_t* xml, const py::object& mapper, const py::object& name)
      {
        wstring addinName;
        if (name.is_none())
        {
          // The returned pointers here do no need to be freed or decref'd
          auto frame = PyEval_GetFrame();
          if (!frame)
            throw py::cast_error();
  #if PY_MAJOR_VERSION >= 4 || PY_MINOR_VERSION >= 9       
          auto code = PyFrame_GetCode(frame);
  #else
          auto code = frame->f_code;
  #endif
          std::filesystem::path filePath(pyToWStr(code->co_filename));
          addinName = filePath.filename().stem();
        }
        else
          addinName = pyToWStr(name.ptr());
        auto addin = makeComAddin(addinName.c_str());
        setRibbon(addin.get(), xml, mapper);
        addin->connect();
        return addin;
      }

      class PyObjectHolder : public pybind11::detail::object_api<PyObjectHolder>
      {
        py::object _obj;
      public:
        PyObjectHolder(const py::object& obj)
          : _obj(obj)
        {}
        ~PyObjectHolder()
        {
          py::gil_scoped_acquire getGil;
          _obj = py::none();
        }
        operator py::object() const { return _obj; }

        /// Return the underlying ``PyObject *`` pointer
        PyObject* ptr() const { return _obj.ptr(); }
        PyObject*& ptr()      { return _obj.ptr(); }
      };

      PyObject* callOneArg(PyObject* callable, PyObject* arg)
      {
#if PY_VERSION_HEX < 0x03080000
        auto result = PyObject_CallFunctionObjArgs(callable, arg, nullptr);
#elseif PY_VERSION_HEX < 0x03090000
          PyObject * args[] = { nullptr, arg };
        auto result = _PyObject_Vectorcall(callable, args + 1, 1 | PY_VECTORCALL_ARGUMENTS_OFFSET, nullptr);
#else
        auto result = PyObject_CallOneArg(callable, arg);
#endif

        return result;
      }

      void addVisibilityCallback(const py::object& self, const py::object& obj)
      {
        auto ctp = self.cast<ICustomTaskPane*>();
        // We take a weak reference to everything - avoid increasing ref count
        // to avoid a circular reference
        // pybind weakref bug https://github.com/pybind/pybind11/issues/2536
        ctp->addVisibilityChangeHandler([
            callable = PyObjectHolder(obj),//py::weakref(static_cast<py::handle>(obj))),
            pane = self.ptr()
          ](ICustomTaskPane&)
          {
            py::gil_scoped_acquire getGil;
            // Get strong ref first then check it, this is thread safe
            PySteal<>(callOneArg(callable.ptr(), pane));
            //auto strong = callable();
            //if (!strong.is_none())
            //  PySteal<>(callOneArg(strong.ptr(), pane));
          });
      }

      void setTaskPaneSize(ICustomTaskPane* pane, const py::object& pair)
      {
        pane->setSize(pair.begin()->cast<int>(), (++pair.begin())->cast<int>());
      }
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<RibbonControl>(mod, "RibbonControl")
          .def_readonly("id", &RibbonControl::Id)
          .def_readonly("tag", &RibbonControl::Tag);

        py::class_<ICustomTaskPane>(mod, "TaskPane")
          .def_property_readonly("hwnd", &ICustomTaskPane::hWnd)
          .def_property("visible", &ICustomTaskPane::getVisible, &ICustomTaskPane::setVisible)
          .def_property("size", &ICustomTaskPane::getSize, setTaskPaneSize)
          .def("visibility_change_callback", addVisibilityCallback, py::arg("callable"));

        py::class_<IComAddin, shared_ptr<IComAddin>>(mod, "ExcelUI")
          .def("connect", &IComAddin::connect)
          .def("disconnect", &IComAddin::disconnect)
          .def("set_ribbon", setRibbon, py::arg("xml"), py::arg("mapper"))
          .def("invalidate", &IComAddin::ribbonInvalidate, py::arg("id") = nullptr)
          .def("activate", &IComAddin::ribbonActivate, py::arg("id"))
          .def("create_task_pane", &IComAddin::createTaskPane, py::arg("name"))
          .def_property_readonly("name", &IComAddin::progid);

        mod.def("create_ribbon", createRibbon, py::arg("xml"), py::arg("mapper"), py::arg("name")=py::none());
      });
    }
  }
}