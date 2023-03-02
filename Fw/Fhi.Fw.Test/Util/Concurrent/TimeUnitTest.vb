Imports Fhi.Fw.Util.Concurrent
Imports NUnit.Framework

Namespace Util.Concurrent
    Public Class TimeUnitTest

        <Test()> Public Sub ToNanos()
            Assert.AreEqual(1&, TimeUnit.NANOSECONDS.ToNanos(1&))
            Assert.AreEqual(1000&, TimeUnit.MICROSECONDS.ToNanos(1&))
            Assert.AreEqual(1000000&, TimeUnit.MILLISECONDS.ToNanos(1&))
            Assert.AreEqual(1000000000&, TimeUnit.SECONDS.ToNanos(1&))
            Assert.AreEqual(1000000000& * 60&, TimeUnit.MINUTES.ToNanos(1&))
            Assert.AreEqual(1000000000& * 60 * 60, TimeUnit.HOURS.ToNanos(1&))
            Assert.AreEqual(1000000000& * 60 * 60 * 24, TimeUnit.DAYS.ToNanos(1&))
        End Sub

        <Test()> Public Sub ToMicros()
            Assert.AreEqual(10&, TimeUnit.NANOSECONDS.ToMicros(10& * 1000&))
            Assert.AreEqual(1&, TimeUnit.MICROSECONDS.ToMicros(1&))
            Assert.AreEqual(1000&, TimeUnit.MILLISECONDS.ToMicros(1&))
            Assert.AreEqual(1000000&, TimeUnit.SECONDS.ToMicros(1&))
            Assert.AreEqual(1000000& * 60&, TimeUnit.MINUTES.ToMicros(1&))
            Assert.AreEqual(1000000& * 60 * 60, TimeUnit.HOURS.ToMicros(1&))
            Assert.AreEqual(1000000& * 60 * 60 * 24, TimeUnit.DAYS.ToMicros(1&))
        End Sub

        <Test()> Public Sub ToMillis()
            Assert.AreEqual(10&, TimeUnit.NANOSECONDS.ToMillis(10& * 1000& * 1000&))
            Assert.AreEqual(10&, TimeUnit.MICROSECONDS.ToMillis(10& * 1000&))
            Assert.AreEqual(1&, TimeUnit.MILLISECONDS.ToMillis(1&))
            Assert.AreEqual(1000&, TimeUnit.SECONDS.ToMillis(1&))
            Assert.AreEqual(1000& * 60&, TimeUnit.MINUTES.ToMillis(1&))
            Assert.AreEqual(1000& * 60 * 60, TimeUnit.HOURS.ToMillis(1&))
            Assert.AreEqual(1000& * 60 * 60 * 24, TimeUnit.DAYS.ToMillis(1&))
        End Sub

        <Test()> Public Sub ToSeconds()
            Assert.AreEqual(10&, TimeUnit.NANOSECONDS.ToSeconds(10& * 1000& * 1000& * 1000&))
            Assert.AreEqual(10&, TimeUnit.MICROSECONDS.ToSeconds(10& * 1000& * 1000&))
            Assert.AreEqual(10&, TimeUnit.MILLISECONDS.ToSeconds(10& * 1000&))
            Assert.AreEqual(1&, TimeUnit.SECONDS.ToSeconds(1&))
            Assert.AreEqual(1& * 60&, TimeUnit.MINUTES.ToSeconds(1&))
            Assert.AreEqual(1& * 60 * 60, TimeUnit.HOURS.ToSeconds(1&))
            Assert.AreEqual(1& * 60 * 60 * 24, TimeUnit.DAYS.ToSeconds(1&))
        End Sub

        <Test()> Public Sub ToMinutes()
            Assert.AreEqual(10&, TimeUnit.NANOSECONDS.ToMinutes(10& * 1000& * 1000& * 1000& * 60&))
            Assert.AreEqual(10&, TimeUnit.MICROSECONDS.ToMinutes(10& * 1000& * 1000& * 60&))
            Assert.AreEqual(10&, TimeUnit.MILLISECONDS.ToMinutes(10& * 1000& * 60&))
            Assert.AreEqual(10&, TimeUnit.SECONDS.ToMinutes(10& * 60&))
            Assert.AreEqual(1&, TimeUnit.MINUTES.ToMinutes(1&))
            Assert.AreEqual(1& * 60, TimeUnit.HOURS.ToMinutes(1&))
            Assert.AreEqual(1& * 60 * 24, TimeUnit.DAYS.ToMinutes(1&))
        End Sub

        <Test()> Public Sub ToHours()
            Assert.AreEqual(10&, TimeUnit.NANOSECONDS.ToHours(10& * 1000& * 1000& * 1000& * 60& * 60&))
            Assert.AreEqual(10&, TimeUnit.MICROSECONDS.ToHours(10& * 1000& * 1000& * 60& * 60&))
            Assert.AreEqual(10&, TimeUnit.MILLISECONDS.ToHours(10& * 1000& * 60& * 60&))
            Assert.AreEqual(10&, TimeUnit.SECONDS.ToHours(10& * 60& * 60&))
            Assert.AreEqual(10&, TimeUnit.MINUTES.ToHours(10& * 60&))
            Assert.AreEqual(1&, TimeUnit.HOURS.ToHours(1&))
            Assert.AreEqual(1& * 24, TimeUnit.DAYS.ToHours(1&))
        End Sub

        <Test()> Public Sub ToDays()
            Assert.AreEqual(10&, TimeUnit.NANOSECONDS.ToDays(10& * 1000& * 1000& * 1000& * 60& * 60& * 24&))
            Assert.AreEqual(10&, TimeUnit.MICROSECONDS.ToDays(10& * 1000& * 1000& * 60& * 60& * 24&))
            Assert.AreEqual(10&, TimeUnit.MILLISECONDS.ToDays(10& * 1000& * 60& * 60& * 24&))
            Assert.AreEqual(10&, TimeUnit.SECONDS.ToDays(10& * 60& * 60& * 24&))
            Assert.AreEqual(10&, TimeUnit.MINUTES.ToDays(10& * 60& * 24&))
            Assert.AreEqual(10&, TimeUnit.HOURS.ToDays(10& * 24&))
            Assert.AreEqual(1&, TimeUnit.DAYS.ToDays(1&))
        End Sub

        <Test()> Public Sub Convert()
            Assert.AreEqual(2&, TimeUnit.NANOSECONDS.Convert(2&, TimeUnit.NANOSECONDS))
            Assert.AreEqual(3& * 1000 * 1000, TimeUnit.MICROSECONDS.Convert(3&, TimeUnit.SECONDS))
            Assert.AreEqual(4& * 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS.Convert(4&, TimeUnit.DAYS))
            Assert.AreEqual(5&, TimeUnit.SECONDS.Convert(5& * 1000, TimeUnit.MILLISECONDS))
            Assert.AreEqual(6& * 60, TimeUnit.MINUTES.Convert(6&, TimeUnit.HOURS))
            Assert.AreEqual(7&, TimeUnit.HOURS.Convert(7& * 1000 * 1000 * 60 * 60, TimeUnit.MICROSECONDS))
            Assert.AreEqual(8&, TimeUnit.DAYS.Convert(8 * 60 * 24, TimeUnit.MINUTES))
        End Sub

    End Class
End Namespace