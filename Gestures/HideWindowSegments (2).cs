using Microsoft.Kinect;
using System.Diagnostics;

namespace PowerPoint_kinect.Gestures
{
    public class HideWindowSegment1 : IRelativeGestureSegment
    {
        public GesturePartResult CheckGesture(Skeleton skeleton)
        {
            // Right and Left Hand in front of Shoulders
            if (skeleton.Joints[JointType.HandLeft].Position.Z < skeleton.Joints[JointType.ShoulderCenter].Position.Z && skeleton.Joints[JointType.HandRight].Position.Z < skeleton.Joints[JointType.ShoulderCenter].Position.Z)
            {
                // Hands very near
                if (skeleton.Joints[JointType.HandRight].Position.X - skeleton.Joints[JointType.HandLeft].Position.X < 0.3)
                {
                    // Hands above head
                    if (skeleton.Joints[JointType.HandRight].Position.Y >= skeleton.Joints[JointType.Head].Position.Y && skeleton.Joints[JointType.HandLeft].Position.Y >= skeleton.Joints[JointType.Head].Position.Y)
                    {
                        //Trace.WriteLine("Passsed 1st segment");
                        return GesturePartResult.Succeed;
                    }

                    return GesturePartResult.Pausing;
                }

                return GesturePartResult.Fail;
            }

            return GesturePartResult.Fail;
        }
    }

    public class HideWindowSegment2 : IRelativeGestureSegment
    {
        public GesturePartResult CheckGesture(Skeleton skeleton)
        {
            // Right and Left Hand in front of Shoulders
            if (skeleton.Joints[JointType.HandLeft].Position.Z < skeleton.Joints[JointType.ShoulderCenter].Position.Z && skeleton.Joints[JointType.HandRight].Position.Z < skeleton.Joints[JointType.ShoulderCenter].Position.Z)
            {
                // Hands very near
                if (skeleton.Joints[JointType.HandRight].Position.X - skeleton.Joints[JointType.HandLeft].Position.X < 0.3)
                {
                    // Hands below head but above SpineCenter
                    if (skeleton.Joints[JointType.HandRight].Position.Y < skeleton.Joints[JointType.Head].Position.Y && skeleton.Joints[JointType.HandRight].Position.Y >= (skeleton.Joints[JointType.Head].Position.Y + skeleton.Joints[JointType.Spine].Position.Y) / 2)
                    {
                        //Trace.WriteLine("Passsed 2nd segment");
                        return GesturePartResult.Succeed;
                    }
                    //Trace.WriteLine("PAUSE - Hands above head or below Spine");
                    return GesturePartResult.Pausing;
                }
               // Trace.WriteLine("FAIL - Hands far");
                return GesturePartResult.Fail;
            }
            return GesturePartResult.Fail;
        }
    }

    public class HideWindowSegment3 : IRelativeGestureSegment
    {
        public GesturePartResult CheckGesture(Skeleton skeleton)
        {

            // Right and Left Hand in front of Shoulders
            if (skeleton.Joints[JointType.HandLeft].Position.Z < skeleton.Joints[JointType.ShoulderCenter].Position.Z && skeleton.Joints[JointType.HandRight].Position.Z < skeleton.Joints[JointType.ShoulderCenter].Position.Z)
            {
                // Hands very near
                if (skeleton.Joints[JointType.HandRight].Position.X - skeleton.Joints[JointType.HandLeft].Position.X < 0.3)
                {
                    // Hands below Spine
                    if (skeleton.Joints[JointType.HandRight].Position.Y <= skeleton.Joints[JointType.Spine].Position.Y && skeleton.Joints[JointType.HandLeft].Position.Y <= skeleton.Joints[JointType.Spine].Position.Y)
                    {
                        return GesturePartResult.Succeed;
                    }
                    return GesturePartResult.Pausing;
                }

                return GesturePartResult.Fail;
            }

            return GesturePartResult.Fail;
        }
    }
}
